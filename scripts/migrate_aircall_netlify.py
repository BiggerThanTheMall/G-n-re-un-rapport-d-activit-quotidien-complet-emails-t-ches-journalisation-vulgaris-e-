from pathlib import Path

P = Path('LTOA-Modulr-Rapport-Quotidien.user.js')
s = P.read_text(encoding='utf-8')

s = s.replace('// @version      5.2.1', '// @version      5.3.0', 1)
s = s.replace('// @match        https://*.aircall.io/*\n', '', 1)
s = s.replace('// @connect      api.aircall.io\n', '', 1)

start = s.index("    const configureAircallApi = () => {")
end_marker = "    GM_registerMenuCommand('Configurer l’API Aircall pour les rapports', configureAircallApi);\n"
end = s.index(end_marker, start) + len(end_marker)
s = s[:start] + s[end:]

start = s.index('    // Le même userscript assure les deux rôles.')
end = s.index('    // ============================================\n    // CONFIGURATION', start)
s = s[:start] + s[end:]

start_marker = "    // ============================================\n    // COLLECTEUR D'APPELS AIRCALL"
end_marker = "    // ============================================\n    // COLLECTEUR DE TÂCHES TERMINÉES"
start = s.index(start_marker)
end = s.index(end_marker, start)

replacement = r'''    // ============================================
    // COLLECTEUR D'APPELS AIRCALL
    // ============================================
    // Accès Aircall sécurisé via Netlify : aucun secret dans Tampermonkey.
    const AircallCollector = {
        lastStatus: { state: 'not_started', source: null, message: '' },
        API_ROOT: 'https://aircallmodulr.netlify.app',

        request(path) {
            return new Promise((resolve, reject) => {
                const controller = new AbortController();
                const timer = setTimeout(() => controller.abort(), 30000);
                fetch(`${this.API_ROOT}${path}`, {
                    method: 'GET', mode: 'cors', credentials: 'omit',
                    headers: { Accept: 'application/json' }, signal: controller.signal
                }).then(async response => {
                    clearTimeout(timer);
                    let body = null;
                    try { body = await response.json(); } catch (_) {}
                    if (response.ok && body) resolve(body);
                    else reject(new Error(body?.error || `Aircall HTTP ${response.status}`));
                }).catch(error => {
                    clearTimeout(timer);
                    reject(new Error(error?.name === 'AbortError' ? 'Délai API Aircall dépassé' : 'Connexion API Aircall impossible'));
                });
            });
        },

        normalizeName(value) {
            return String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
                .toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
        },

        async findUserId(userName) {
            const aliases = {
                'Doryan KALAH': 'Doryan Kalah', 'Eddy KALAH': 'Eddy Kalah',
                'Ghais Kalah': 'Ghais Kalah', 'GHAIS KALAH': 'Ghais Kalah',
                'Jake CASIMIR': 'Jake CASIMIR', 'Louli VULLIOD-PIN': 'Louli VULLIOD',
                'Nadia KALAH': 'Nadia Kalah', 'Youness OUACHBAB': 'Youness OUACHBAB',
                'Sheana KRIEF': 'Sheana KRIEF'
            };
            const wanted = this.normalizeName(aliases[userName] || userName);
            for (let page = 1; page <= 250; page++) {
                const payload = await this.request(`/api/report/users?per_page=50&page=${page}`);
                const user = (payload.users || []).find(item => this.normalizeName(item.name) === wanted);
                if (user) return user.id;
                if (!payload.meta?.next_page_link) break;
            }
            throw new Error(`Collaborateur Aircall introuvable : ${userName}`);
        },

        dateRange(dateStr) {
            const [day, month, year] = String(dateStr).split('/').map(Number);
            if (!day || !month || !year) throw new Error('Date Aircall invalide');
            return {
                from: Math.floor(new Date(year, month - 1, day, 0, 0, 0).getTime() / 1000),
                to: Math.floor(new Date(year, month - 1, day + 1, 0, 0, 0).getTime() / 1000) - 1
            };
        },

        async insight(callId, endpoint) {
            await Utils.delay(650);
            try {
                return await this.request(`/api/report/calls/${encodeURIComponent(callId)}/${encodeURIComponent(endpoint)}`);
            } catch (error) {
                Utils.log(`Aircall ${endpoint} indisponible pour ${callId}: ${error.message}`);
                return null;
            }
        },

        async collectDirect(connectedUser, reportDate, updateLoader) {
            const userId = await this.findUserId(connectedUser);
            const range = this.dateRange(reportDate);
            const calls = [], seen = new Set();
            for (let page = 1; page <= 250; page++) {
                updateLoader(`API Aircall via Netlify : page ${page}...`);
                const q = new URLSearchParams({
                    user_id: String(userId), from: String(range.from), to: String(range.to),
                    order: 'asc', per_page: '50', fetch_contact: 'true', page: String(page)
                });
                const payload = await this.request(`/api/report/calls?${q.toString()}`);
                for (const call of payload.calls || []) {
                    if (seen.has(call.id)) continue;
                    seen.add(call.id);
                    const contact = call.contact?.first_name || call.contact?.last_name
                        ? [call.contact.first_name, call.contact.last_name].filter(Boolean).join(' ')
                        : (call.contact?.name || call.raw_digits || 'Inconnu');
                    calls.push({
                        id: call.id,
                        type: call.direction === 'inbound' ? 'entrant' : 'sortant',
                        user: call.user?.name || connectedUser,
                        contact,
                        phone: call.raw_digits || '',
                        durationSeconds: Number(call.duration) || 0,
                        duration: `${Math.floor((Number(call.duration) || 0) / 60)}m ${(Number(call.duration) || 0) % 60}s`,
                        time: call.started_at ? new Date(call.started_at * 1000).toLocaleTimeString('fr-FR', {hour:'2-digit', minute:'2-digit'}) : '',
                        answered: Boolean(call.answered_at),
                        missedReason: call.missed_call_reason || null,
                        tags: (call.tags || []).map(tag => tag.name).filter(Boolean),
                        summary: (call.comments || []).map(item => item.content).filter(Boolean).join(' — ') || null,
                        source: 'api'
                    });
                }
                if (!payload.meta?.next_page_link) break;
            }

            const answered = calls.filter(call => call.answered);
            for (let index = 0; index < answered.length; index++) {
                const call = answered[index];
                updateLoader(`Analyses Aircall ${index + 1}/${answered.length}...`);
                const summaryPayload = await this.insight(call.id, 'summary');
                const sentimentPayload = await this.insight(call.id, 'sentiments');
                const topicsPayload = await this.insight(call.id, 'topics');
                const actionsPayload = await this.insight(call.id, 'action_items');
                const transcriptPayload = await this.insight(call.id, 'transcription');
                const externalSentiment = sentimentPayload?.sentiment?.participants?.find(p => p.type === 'external')?.value
                    || sentimentPayload?.sentiment?.participants?.[0]?.value || null;
                const moodMap = { POSITIVE: 'Positif', NEGATIVE: 'Négatif', NEUTRAL: 'Neutre' };
                const utterances = transcriptPayload?.transcription?.content?.utterances || [];
                call.summary = summaryPayload?.summary?.content || call.summary || null;
                call.mood = moodMap[String(externalSentiment || '').toUpperCase()] || externalSentiment || null;
                call.topics = Array.isArray(topicsPayload?.topic?.content) ? topicsPayload.topic.content : [];
                call.actionItems = (actionsPayload?.action_items || []).map(item => typeof item === 'string' ? item : item?.content).filter(Boolean);
                call.transcript = utterances.map(item => {
                    const speaker = item.participant_type === 'internal' ? 'Collaborateur' : (item.participant_type === 'external' ? 'Interlocuteur' : 'Autre');
                    return `${speaker} : ${item.text || ''}`;
                }).filter(line => !line.endsWith(': ')).join('\n');
                call.transcriptLanguage = transcriptPayload?.transcription?.content?.language || null;
            }
            return calls;
        },

        async collect(connectedUser, reportDate, updateLoader) {
            if (!CONFIG.AIRCALL_ENABLED) {
                this.lastStatus = { state: 'disabled', source: null, message: 'Aircall désactivé' };
                return [];
            }
            try {
                updateLoader('Connexion Aircall sécurisée via Netlify...');
                const calls = await this.collectDirect(connectedUser, reportDate, updateLoader);
                this.lastStatus = { state: 'complete', source: 'api', message: `API Aircall via Netlify : ${calls.length} appel${calls.length > 1 ? 's' : ''} trouvé${calls.length > 1 ? 's' : ''}` };
                return calls;
            } catch (error) {
                this.lastStatus = { state: 'error', source: 'api', message: `Erreur API Aircall : ${error.message}` };
                throw error;
            }
        }
    };

'''

s = s[:start] + replacement + s[end:]

forbidden = ['ltoa_aircall_api_id', 'ltoa_aircall_api_token', 'api.aircall.io', 'configureAircallApi', 'Authorization: `Basic']
left = [x for x in forbidden if x in s]
if left:
    for n, line in enumerate(s.splitlines(), 1):
        if any(x in line for x in left):
            print(f'RESIDUAL {n}: {line[:220]}')
    raise SystemExit(f'Forbidden Aircall references remain: {left}')
if 'https://aircallmodulr.netlify.app' not in s:
    raise SystemExit('Netlify root missing')

P.write_text(s, encoding='utf-8')
print('Migration applied successfully')
