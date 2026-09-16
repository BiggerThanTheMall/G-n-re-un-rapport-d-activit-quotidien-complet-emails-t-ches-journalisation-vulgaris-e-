// ==UserScript==
// @name         LTOA Modulr - Rapport Quotidien
// @namespace    https://github.com/BiggerThanTheMall/tampermonkey-ltoa
// @version      5.3.0
// @description  Génération automatique du rapport d’activité quotidien dans Modulr
// @author       LTOA Assurances
// @match        https://courtage.modulr.fr/*
// @exclude      https://courtage.modulr.fr/fr/intranet/edm/preview/document/*
// @exclude      https://courtage.modulr.fr/fr/intranet/edm/display/Client/*
// @exclude      https://courtage.modulr.fr/fr/scripts/sent_emails/sent_emails_frame.php?sent_email_id*
// @run-at       document-end
// @grant        GM_xmlhttpRequest
// @grant        GM_openInTab
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_registerMenuCommand
// @connect      courtage.modulr.fr
//
// @updateURL    https://raw.githubusercontent.com/BiggerThanTheMall/G-n-re-un-rapport-d-activit-quotidien-complet-emails-t-ches-journalisation-vulgaris-e-/main/LTOA-Modulr-Rapport-Quotidien.user.js
// @downloadURL  https://raw.githubusercontent.com/BiggerThanTheMall/G-n-re-un-rapport-d-activit-quotidien-complet-emails-t-ches-journalisation-vulgaris-e-/main/LTOA-Modulr-Rapport-Quotidien.user.js
// ==/UserScript==

(function() {
    'use strict';


    // ============================================
    // CONFIGURATION
    // ============================================
    const CONFIG = {
        // Pour créer le webhook Teams :
        // 1. Ouvrir Teams > Canal "GENERAL LTOA"
        // 2. Clic droit sur le canal > "Connecteurs"
        // 3. Chercher "Incoming Webhook" > Configurer
        // 4. Donner un nom (ex: "Rapport LTOA") > Créer
        // 5. Copier l'URL générée et la coller ici :
        TEAMS_WEBHOOK_URL: '',

        DEBUG: true,
        DELAY_BETWEEN_REQUESTS: 600,
        DELAY_EMAIL_BODY: 800,
        MAX_PAGES_TO_CHECK: 10,

        // Configuration Aircall
        AIRCALL_ENABLED: true,
        AIRCALL_TIMEOUT: 900000, // jusqu'à 15 minutes si de nombreux appels sont enrichis
    };

    // ============================================
    // DATE SÉLECTIONNÉE POUR LE RAPPORT
    // ============================================
    let SELECTED_REPORT_DATE = null; // Format: DD/MM/YYYY ou null pour aujourd'hui
    let REPORT_NOTES = '';

    // ============================================
    // MAPPING DES UTILISATEURS
    // ============================================
    const USER_MAP = {
        'Doryan KALAH': { taskValue: 'user_id:33', logValue: '33', emailFilter: 'Doryan KALAH' },
        'Eddy KALAH': { taskValue: 'user_id:23', logValue: '23', emailFilter: 'Eddy KALAH' },
        'Ghais Kalah': { taskValue: 'user_id:24', logValue: '24', emailFilter: 'Ghais Kalah' },
        'GHAIS KALAH': { taskValue: 'user_id:24', logValue: '24', emailFilter: 'Ghais Kalah' },
        'Jake CASIMIR': { taskValue: 'user_id:28', logValue: '28', emailFilter: 'Jake CASIMIR' },
        'Louli VULLIOD-PIN': { taskValue: 'user_id:36', logValue: '36', emailFilter: 'Louli VULLIOD-PIN' },
        'Nadia KALAH': { taskValue: 'user_id:22', logValue: '22', emailFilter: 'Nadia KALAH' },
        'Sheana KRIEF': { taskValue: 'user_id:2', logValue: '2', emailFilter: 'Sheana KRIEF' },
        'Youness OUACHBAB': { taskValue: 'user_id:39', logValue: '39', emailFilter: 'Youness OUACHBAB' },
        'Faicel BEN LASWED': { taskValue: 'user_id:32', logValue: '32', emailFilter: 'Faicel BEN LASWED' },
        'Inssaf CHOUAOUA': { taskValue: 'user_id:25', logValue: '25', emailFilter: 'Inssaf CHOUAOUA' },
        'Wesley DAUX': { taskValue: 'user_id:29', logValue: '29', emailFilter: 'Wesley DAUX' },
    };

    // ============================================
    // TRADUCTIONS POUR VULGARISATION
    // ============================================
    const TRANSLATIONS = {
        // Noms de tables
        tables: {
            'Tâches': 'Tâches',
            'Tasks': 'Tâches',
            'Emails envoyés': 'Emails',
            'sent_emails': 'Emails',
            'Clients': 'Clients',
            'clients': 'Clients',
            'Prospects': 'Prospects',
            'Contrats': 'Contrats',
            'contracts': 'Contrats',
            'Sinistres': 'Sinistres',
            'claims': 'Sinistres',
            'Devis': 'Devis',
            'estimates': 'Devis',
        },

        // Noms de champs techniques -> noms lisibles
        fields: {
            // Identité
            'name': 'Nom',
            'first_name': 'Prénom',
            'firstname': 'Prénom',
            'last_name': 'Nom',
            'lastname': 'Nom',
            'title': 'Civilité',
            'civility': 'Civilité',
            'birth_date': 'Date de naissance',
            'birthdate': 'Date de naissance',
            'birth_country': 'Pays de naissance',
            'birth_location': 'Lieu de naissance',
            'birth_place': 'Lieu de naissance',
            'nationality': 'Nationalité',

            // Coordonnées
            'email': 'Email',
            'phone': 'Téléphone',
            'phone_1': 'Téléphone 1',
            'phone_2': 'Téléphone 2',
            'mobile': 'Mobile',
            'mobile_phone': 'Téléphone mobile',
            'fax': 'Fax',
            'address': 'Adresse',
            'address_1': 'Adresse',
            'address_2': 'Complément adresse',
            'postal_code': 'Code postal',
            'zip_code': 'Code postal',
            'city': 'Ville',
            'country': 'Pays',

            // Coordonnées bancaires
            'iban': 'IBAN',
            'bic': 'BIC',
            'bank_name': 'Banque',
            'bank_domiciliation': 'Domiciliation bancaire',
            'bank_account_holder': 'Titulaire du compte',

            // Statuts et dates système
            'status': 'Statut',
            'client_status': 'Statut client',
            'creation_date': 'Date de création',
            'last_update': 'Dernière modification',
            'last_update_user_id': 'Modifié par (ID)',
            'creation_user_id': 'Créé par (ID)',

            // Devis
            'estimate_id': 'N° Devis',
            'input_date': 'Date de saisie',
            'validity_date': 'Date de validité',
            'expiry_date': 'Date d\'expiration',
            'expiration_date': 'Date d\'expiration',
            'product_type_id': 'Type de produit',
            'product_id': 'Produit',
            'company_id': 'Compagnie',
            'premium': 'Prime',
            'total_amount': 'Montant total',
            'commission': 'Commission',
            'office_id': 'Bureau',
            'firm_id': 'Cabinet',
            'bank_account_id': 'Compte bancaire',
            'client_id': 'Client (ID)',
            'referent_user_id': 'Gestionnaire référent',
            'client_communications_recipient': 'Destinataire communications',

            // Contrats
            'policy_id': 'N° Contrat',
            'ref': 'Référence',
            'reference': 'Référence',
            'effective_date': 'Date d\'effet',
            'start_date': 'Date de début',
            'end_date': 'Date de fin',
            'renewal_date': 'Date de renouvellement',
            'expiration_date': 'Date d\'expiration',
            'expiration_detail': 'Détail expiration',
            'displayed_in_extranet': 'Visible sur extranet',
            'beneficiaries': 'Bénéficiaires',
            'end_date_annual_declaration': 'Fin déclaration annuelle',
            'deducted_commissions': 'Commissions déduites',
            'business_type': 'Type d\'affaire',
            'application_fee_calculation_source': 'Source calcul frais',
            'application_fee_on_premium_per_type': 'Frais sur prime par type',
            'claim_payment_external': 'Paiement sinistre externe',
            'update_guarantee_from_index': 'MAJ garantie depuis index',

            // Sinistres
            'claim_id': 'N° Sinistre',
            'claim_date': 'Date du sinistre',
            'declaration_date': 'Date de déclaration',
            'closing_date': 'Date de clôture',
            'trouble_ticket': 'N° Dossier',
            'client_reference': 'Référence client',
            'guarantee_id': 'Garantie',
            'comment': 'Commentaire',

            // Tâches
            'task_id': 'N° Tâche',
            'task_type': 'Type de tâche',
            'event_type': 'Type d\'événement',
            'due_date': 'Date d\'échéance',
            'priority': 'Priorité',
            'recipient': 'Destinataire',
            'creator': 'Créateur',
            'description': 'Description',
            'content': 'Contenu',
            'origin': 'Origine',
            'notes': 'Notes',

            // Emails
            'subject': 'Objet',
            'body': 'Contenu',
            'to': 'Destinataire',
            'from': 'Expéditeur',
            'cc': 'Copie',
            'bcc': 'Copie cachée',
            'attachments': 'Pièces jointes',
            'email_origin': 'Origine de l\'email',
        },

        // Valeurs de champs -> valeurs lisibles
        values: {
            // Booléens
            'yes': 'Oui',
            'no': 'Non',
            'true': 'Oui',
            'false': 'Non',
            '1': 'Oui',
            '0': 'Non',

            // Statuts devis
            'current': 'En cours',
            'pricing': 'En tarification',
            'delivered': 'Transmis au client',
            'accepted': 'Accepté',
            'refused': 'Refusé',
            'expired': 'Expiré',
            'cancelled': 'Annulé',
            'waiting': 'En attente',
            'validated': 'Validé',

            // Statuts contrat
            'active': 'Actif',
            'inactive': 'Inactif',
            'suspended': 'Suspendu',
            'terminated': 'Résilié',
            'renewed': 'Renouvelé',
            'in_force': 'En vigueur',
            '10': 'En vigueur',

            // Statuts sinistre
            'open': 'Ouvert',
            'closed': 'Clôturé',
            'in_progress': 'En cours',
            '4': 'En cours de traitement',

            // Destinataires communications
            'client': 'Client',
            'producer': 'Apporteur',
            'manager': 'Gestionnaire',

            // Priorités
            'high': 'Haute',
            'normal': 'Normale',
            'low': 'Basse',

            // Statuts tâche
            'pending': 'En attente',
            'finished': 'Terminée',

            // Pays
            'FRANCE': 'France',
            'France': 'France',

            // Origines
            'automatic': 'Automatique',
            'manual': 'Manuel',
            'system': 'Système',
        }
    };

    // ============================================
    // VULGARISATEUR DE LOGS
    // ============================================
    const LogVulgarizer = {
        // Générer un résumé vulgarisé d'une entrée de log
        vulgarize(entry) {
            const action = entry.actionRaw || entry.action;
            const table = entry.table || entry.tableRaw;
            const entityName = entry.entityName || '';
            const changes = entry.changes || [];

            // Déterminer l'icône et le verbe selon l'action
            let icon = '📝';
            let verb = '';

            if (action.includes('Insertion')) {
                icon = '✨';
                verb = this.getCreationVerb(table);
            } else if (action.includes('Mise à jour')) {
                icon = '✏️';
                verb = this.getUpdateVerb(table);
            } else if (action.includes('Suppression')) {
                icon = '🗑️';
                verb = this.getDeleteVerb(table);
            }

            // Construire le titre vulgarisé
            let title = `${icon} ${verb}`;
            if (entityName && entityName !== 'N/A') {
                title += ` : ${entityName}`;
            }

            // Résumer les changements importants
            const summary = this.summarizeChanges(changes, table, action);

            return {
                icon,
                title,
                summary,
                details: this.formatChangesForDisplay(changes)
            };
        },

        getCreationVerb(table) {
            const verbs = {
                'Clients': 'Nouveau client créé',
                'Client': 'Nouveau client créé',
                'Devis': 'Nouveau devis créé',
                'Contrats': 'Nouveau contrat souscrit',
                'Sinistres': 'Nouveau sinistre déclaré',
            };
            return verbs[table] || `Création ${table}`;
        },

        getUpdateVerb(table) {
            const verbs = {
                'Clients': 'Fiche client modifiée',
                'Client': 'Fiche client modifiée',
                'Devis': 'Devis mis à jour',
                'Contrats': 'Contrat modifié',
                'Sinistres': 'Sinistre mis à jour',
            };
            return verbs[table] || `Mise à jour ${table}`;
        },

        getDeleteVerb(table) {
            const verbs = {
                'Clients': 'Client supprimé',
                'Devis': 'Devis supprimé',
                'Contrats': 'Contrat supprimé',
                'Sinistres': 'Sinistre supprimé',
            };
            return verbs[table] || `Suppression ${table}`;
        },

        summarizeChanges(changes, table, action) {
            if (!changes || changes.length === 0) {
                if (action.includes('Insertion')) {
                    return 'Nouvelle entrée créée';
                }
                return '';
            }

            // Filtrer les champs système
            const systemFields = ['last_update', 'last_update_user_id', 'creation_date', 'creation_user_id', 'id'];
            const meaningfulChanges = changes.filter(c => !systemFields.includes(c.fieldRaw));

            if (meaningfulChanges.length === 0) return '';

            // Générer un résumé intelligent
            const summaryParts = [];
            const processedCategories = new Set();

            for (const change of meaningfulChanges) {
                const field = change.fieldRaw;
                const newVal = change.newValueRaw || change.newValue || '';
                const oldVal = change.oldValueRaw || change.oldValue || '';

                // Grouper par catégorie pour éviter répétitions
                if ((field.includes('address') || field === 'city' || field === 'postal_code' || field === 'country') && !processedCategories.has('address')) {
                    summaryParts.push('📍 Adresse modifiée');
                    processedCategories.add('address');
                } else if ((field.includes('iban') || field.includes('bic') || field.includes('bank')) && !processedCategories.has('bank')) {
                    summaryParts.push('🏦 Coordonnées bancaires');
                    processedCategories.add('bank');
                } else if (field === 'status') {
                    const translatedNew = Utils.translateValue(newVal);
                    summaryParts.push(`📊 Statut → ${translatedNew}`);
                } else if (field === 'comment' && !processedCategories.has('comment')) {
                    summaryParts.push('💬 Commentaire ajouté');
                    processedCategories.add('comment');
                } else if (!processedCategories.has(field) && summaryParts.length < 3) {
                    const fieldName = Utils.translateField(field);
                    if (oldVal === '-' || oldVal === '' || String(oldVal).startsWith('Taille')) {
                        summaryParts.push(`${fieldName} renseigné`);
                    } else {
                        summaryParts.push(`${fieldName} modifié`);
                    }
                    processedCategories.add(field);
                }
            }

            // Limiter et indiquer si plus de changements
            if (meaningfulChanges.length > 3 && summaryParts.length >= 3) {
                return summaryParts.slice(0, 2).join(' • ') + ` (+${meaningfulChanges.length - 2} autres)`;
            }
            return summaryParts.join(' • ');
        },

        formatChangesForDisplay(changes) {
            if (!changes || changes.length === 0) return [];

            const systemFields = ['last_update', 'last_update_user_id', 'creation_date', 'creation_user_id'];

            return changes
                .filter(c => !systemFields.includes(c.fieldRaw))
                .map(c => ({
                    field: Utils.translateField(c.fieldRaw),
                    oldValue: Utils.translateValue(c.oldValueRaw || c.oldValue),
                    newValue: Utils.translateValue(c.newValueRaw || c.newValue)
                }));
        }
    };


    // ============================================
    // UTILITAIRES
    // ============================================
    const Utils = {
        log: (msg, data = null) => {
            if (CONFIG.DEBUG) {
                console.log(`[LTOA-Report] ${msg}`, data || '');
            }
        },

        // Retourne la date du rapport (sélectionnée ou aujourd'hui)
        getTodayDate: () => {
            // Si une date est sélectionnée, l'utiliser
            if (SELECTED_REPORT_DATE) {
                return SELECTED_REPORT_DATE;
            }
            // Sinon, date du jour
            const today = new Date();
            const day = String(today.getDate()).padStart(2, '0');
            const month = String(today.getMonth() + 1).padStart(2, '0');
            const year = today.getFullYear();
            return `${day}/${month}/${year}`;
        },

        // Retourne la date réelle d'aujourd'hui (pour comparaisons)
        getRealTodayDate: () => {
            const today = new Date();
            const day = String(today.getDate()).padStart(2, '0');
            const month = String(today.getMonth() + 1).padStart(2, '0');
            const year = today.getFullYear();
            return `${day}/${month}/${year}`;
        },

        // Retourne J-1 par rapport à la date du rapport
        getYesterdayFromReportDate: () => {
            let baseDate;
            if (SELECTED_REPORT_DATE) {
                // Parser la date sélectionnée DD/MM/YYYY
                const parts = SELECTED_REPORT_DATE.split('/');
                baseDate = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
            } else {
                baseDate = new Date();
            }
            baseDate.setDate(baseDate.getDate() - 1);
            const day = String(baseDate.getDate()).padStart(2, '0');
            const month = String(baseDate.getMonth() + 1).padStart(2, '0');
            const year = baseDate.getFullYear();
            return `${day}/${month}/${year}`;
        },

        // Convertir DD/MM/YYYY en objet Date
        parseDate: (dateStr) => {
            if (!dateStr) return null;
            const parts = dateStr.split('/');
            if (parts.length !== 3) return null;
            return new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
        },

        // Nettoyer le texte (enlever les \r\n\t, balises HTML, entités, caractères spéciaux)
        cleanText: (text) => {
            if (!text) return '';

            let result = text;

            // Étape 1: Convertir les balises de saut de ligne en marqueur temporaire
            result = result.replace(/<br\s*\/?>/gi, '[[NEWLINE]]');
            result = result.replace(/<\/p>/gi, '[[NEWLINE]]');
            result = result.replace(/<\/div>/gi, '[[NEWLINE]]');
            result = result.replace(/<\/li>/gi, '[[NEWLINE]]');

            // Étape 2: Supprimer toutes les autres balises HTML
            result = result.replace(/<[^>]+>/g, '');

            // Étape 3: Décoder les entités HTML
            result = result.replace(/&nbsp;/g, ' ');
            result = result.replace(/&amp;/g, '&');
            result = result.replace(/&lt;/g, '<');
            result = result.replace(/&gt;/g, '>');
            result = result.replace(/&quot;/g, '"');
            result = result.replace(/&#0*39;/g, "'"); // &#39; ou &#039;
            result = result.replace(/&apos;/g, "'");
            result = result.replace(/&#0*34;/g, '"'); // &#34; ou &#034;
            result = result.replace(/&#(\d+);/g, (match, dec) => String.fromCharCode(parseInt(dec)));
            result = result.replace(/&#x([0-9a-f]+);/gi, (match, hex) => String.fromCharCode(parseInt(hex, 16)));

            // Étape 4: Nettoyer les séquences d'échappement littérales (comme dans le texte "\n")
            // Ces patterns apparaissent quand le texte contient littéralement \n, \r, \t
            result = result.replace(/\\r\\n/g, '[[NEWLINE]]');
            result = result.replace(/\\n\\r/g, '[[NEWLINE]]');
            result = result.replace(/\\r/g, '[[NEWLINE]]');
            result = result.replace(/\\n/g, '[[NEWLINE]]');
            result = result.replace(/\\t/g, ' ');

            // Étape 5: Nettoyer les vrais caractères de contrôle
            result = result.replace(/\r\n/g, '[[NEWLINE]]');
            result = result.replace(/\n\r/g, '[[NEWLINE]]');
            result = result.replace(/\r/g, '[[NEWLINE]]');
            result = result.replace(/\n/g, '[[NEWLINE]]');
            result = result.replace(/\t/g, ' ');

            // Étape 6: Supprimer les caractères de contrôle et non-imprimables
            result = result.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '');

            // Étape 7: Restaurer les sauts de ligne
            result = result.replace(/\[\[NEWLINE\]\]/g, '\n');

            // Étape 8: Normaliser les espaces et sauts de ligne
            result = result.replace(/\n{3,}/g, '\n\n');  // Max 2 sauts de ligne consécutifs
            result = result.replace(/[ \t]+/g, ' ');     // Espaces multiples -> 1 espace
            result = result.replace(/^ +/gm, '');        // Espaces en début de ligne
            result = result.replace(/ +$/gm, '');        // Espaces en fin de ligne
            result = result.replace(/\n +\n/g, '\n\n');  // Lignes avec seulement des espaces

            return result.trim();
        },

        getConnectedUser: () => {
            const users = Object.keys(USER_MAP);

            // MÉTHODE PRINCIPALE: div.connectedUser contient le span.tooltip avec le nom
            const connectedUserDiv = document.querySelector('.connectedUser span.tooltip');
            if (connectedUserDiv) {
                const title = connectedUserDiv.getAttribute('title') || '';
                const text = connectedUserDiv.textContent.trim();
                const nameToCheck = title || text;

                for (const user of users) {
                    if (nameToCheck.toLowerCase().includes(user.toLowerCase()) ||
                        user.toLowerCase().includes(nameToCheck.toLowerCase())) {
                        Utils.log('Utilisateur détecté (.connectedUser):', user);
                        return user;
                    }
                }
            }

            // FALLBACK 1: span.tooltip avec fa-user
            const userSpan = document.querySelector('span.tooltip span.fa-user');
            if (userSpan && userSpan.parentElement) {
                const parentText = userSpan.parentElement.textContent.trim();
                const oldTitle = userSpan.parentElement.getAttribute('oldtitle') || '';

                for (const user of users) {
                    if (oldTitle.toLowerCase().includes(user.toLowerCase()) ||
                        parentText.toLowerCase().includes(user.toLowerCase())) {
                        Utils.log('Utilisateur détecté (fa-user):', user);
                        return user;
                    }
                }
            }

            // DERNIER RECOURS: Demander
            Utils.log('Utilisateur non détecté, demande manuelle');
            const userList = users.filter((u, i, arr) => arr.findIndex(x => x.toLowerCase() === u.toLowerCase()) === i).join('\n');
            const choice = prompt(`Utilisateur non détecté.\n\nQui êtes-vous ?\n${userList}`);
            if (choice) {
                for (const user of users) {
                    if (user.toLowerCase().includes(choice.toLowerCase()) ||
                        choice.toLowerCase().includes(user.split(' ')[0].toLowerCase())) {
                        Utils.log('Utilisateur choisi:', user);
                        return user;
                    }
                }
            }

            return 'Utilisateur inconnu';
        },

        getUserData: (name) => {
            const normalizedName = Object.keys(USER_MAP).find(key =>
                key.toLowerCase() === name.toLowerCase()
            );
            return USER_MAP[normalizedName] || USER_MAP['Ghais Kalah'];
        },

        delay: (ms) => new Promise(resolve => setTimeout(resolve, ms)),

        fetchPage: (url) => {
            return new Promise((resolve, reject) => {
                GM_xmlhttpRequest({
                    method: 'GET',
                    url: url,
                    onload: (response) => {
                        if (response.status === 200) {
                            resolve(response.responseText);
                        } else {
                            reject(new Error(`HTTP ${response.status}`));
                        }
                    },
                    onerror: (error) => reject(error)
                });
            });
        },

        // Requête POST (pour les formulaires comme UsersLogsList)
        fetchPagePost: (url, data) => {
            return new Promise((resolve, reject) => {
                GM_xmlhttpRequest({
                    method: 'POST',
                    url: url,
                    headers: {
                        'Content-Type': 'application/x-www-form-urlencoded'
                    },
                    data: data,
                    onload: (response) => {
                        if (response.status === 200) {
                            resolve(response.responseText);
                        } else {
                            reject(new Error(`HTTP ${response.status}`));
                        }
                    },
                    onerror: (error) => reject(error)
                });
            });
        },

        parseHTML: (html) => {
            const parser = new DOMParser();
            return parser.parseFromString(html, 'text/html');
        },

        // Traduire un nom de champ
        translateField: (field) => {
            return TRANSLATIONS.fields[field] || field;
        },

        // Traduire une valeur
        translateValue: (value) => {
            if (value === null || value === undefined || value === '-') return '-';
            let strValue = String(value).trim();
            // Nettoyer les balises HTML si présentes
            if (strValue.includes('<') && strValue.includes('>')) {
                strValue = Utils.cleanText(strValue);
            }
            // Décoder les entités HTML
            strValue = strValue.replace(/&#0*39;/g, "'").replace(/&#0*34;/g, '"').replace(/&amp;/g, '&');
            return TRANSLATIONS.values[strValue] || strValue;
        },

        // Traduire un nom de table
        translateTable: (table) => {
            return TRANSLATIONS.tables[table] || table;
        },

        // Tronquer un texte
        truncate: (text, maxLength = 100) => {
            if (!text) return '';
            if (text.length <= maxLength) return text;
            return text.substring(0, maxLength) + '...';
        },

        // Échapper HTML
        escapeHtml: (text) => {
            if (!text) return '';
            const div = document.createElement('div');
            div.textContent = text;
            return div.innerHTML;
        }
    };

    // ============================================
    // CACHE DES NOMS DE CLIENTS
    // ============================================
    const ClientCache = {
        cache: {},

        async getClientName(clientId) {
            if (!clientId || clientId === '-') return null;

            // Vérifier le cache
            if (this.cache[clientId]) {
                return this.cache[clientId];
            }

            try {
                const url = `https://courtage.modulr.fr/fr/scripts/clients/clients_card.php?id=${clientId}`;
                const html = await Utils.fetchPage(url);
                const doc = Utils.parseHTML(html);

                // Chercher le nom du client dans la page
                // Généralement dans un h1 ou un élément avec le nom
                let name = null;

                // Essayer différents sélecteurs
                const nameSelectors = [
                    'h1',
                    '.client_name',
                    '.entity_name',
                    '#client_name',
                    'span.font_size_higher'
                ];

                for (const selector of nameSelectors) {
                    const el = doc.querySelector(selector);
                    if (el && el.textContent.trim()) {
                        name = el.textContent.trim();
                        break;
                    }
                }

                // Alternative: chercher prénom + nom
                if (!name) {
                    const firstName = doc.querySelector('input[name="first_name"], #first_name');
                    const lastName = doc.querySelector('input[name="last_name"], #last_name');
                    const companyName = doc.querySelector('input[name="company_name"], #company_name');

                    if (companyName && companyName.value) {
                        name = companyName.value;
                    } else if (firstName || lastName) {
                        name = [
                            firstName?.value || '',
                            lastName?.value || ''
                        ].filter(Boolean).join(' ');
                    }
                }

                if (name) {
                    this.cache[clientId] = name;
                    return name;
                }

                return null;
            } catch (error) {
                Utils.log(`Erreur récupération client ${clientId}:`, error);
                return null;
            }
        }
    };

    // ============================================
    // COLLECTEUR D'EMAILS ENVOYÉS
    // ============================================
    const EmailsSentCollector = {
        async collect(connectedUser, updateLoader) {
            Utils.log('Collecte des emails envoyés...');
            const results = [];
            const reportDate = Utils.getTodayDate(); // Date du rapport (peut être dans le passé)
            const reportDateObj = Utils.parseDate(reportDate);

            let currentPage = 1;
            let hasMorePages = true;
            let emailCount = 0;
            let foundReportDateEmails = false;
            let passedReportDate = false; // True quand on a dépassé la date du rapport (emails plus anciens)

            try {
                while (hasMorePages && currentPage <= CONFIG.MAX_PAGES_TO_CHECK && !passedReportDate) {
                    updateLoader(`Emails envoyés - Page ${currentPage}...`);

                    // URL des emails envoyés
                    const url = `https://courtage.modulr.fr/fr/scripts/emails/emails_list.php?sent_email_page=${currentPage}#entity_menu_emails=1`;
                    const html = await Utils.fetchPage(url);
                    const doc = Utils.parseHTML(html);

                    // Les lignes principales sont s_main_XXXX (pas e_main_)
                    const emailRows = doc.querySelectorAll('tr[id^="s_main_"]');

                    Utils.log(`Page ${currentPage}: ${emailRows.length} emails trouvés`);

                    if (emailRows.length === 0) {
                        Utils.log('Aucun email trouvé, fin de la collecte');
                        hasMorePages = false;
                        break;
                    }

                    for (const row of emailRows) {
                        // Récupérer toutes les cellules td avec data-sent_email_id
                        const cells = row.querySelectorAll('td[data-sent_email_id]');
                        if (cells.length < 3) continue;

                        // 1ère cellule = Date
                        const dateCell = cells[0];
                        const dateSpan = dateCell.querySelector('span.middle_fade');
                        const dateText = dateSpan ? dateSpan.textContent.trim() : '';

                        // Extraire la date au format DD/MM/YYYY
                        const dateMatch = dateText.match(/(\d{2}\/\d{2}\/\d{4})/);
                        const emailDate = dateMatch ? dateMatch[1] : '';
                        const emailDateObj = Utils.parseDate(emailDate);

                        // Extraire l'heure au format HH:MM
                        const timeMatch = dateText.match(/(\d{1,2}):(\d{2})/);
                        const emailTime = timeMatch ? `${timeMatch[1].padStart(2, '0')}:${timeMatch[2]}` : '';

                        Utils.log(`Email date: "${emailDate}", reportDate: "${reportDate}"`);

                        // Comparer les dates
                        if (emailDateObj && reportDateObj) {
                            // Si l'email est APRÈS la date du rapport → continuer (pas encore arrivé)
                            if (emailDateObj > reportDateObj) {
                                Utils.log(`  Email plus récent que ${reportDate}, on continue...`);
                                continue;
                            }

                            // Si l'email est à la date du rapport → collecter
                            if (emailDate === reportDate) {
                                foundReportDateEmails = true;
                                emailCount++;

                                // ID de l'email
                                const emailId = dateCell.getAttribute('data-sent_email_id');

                                // 3ème cellule = Destinataire (index 2)
                                const toCell = cells[2];
                                const toSpan = toCell.querySelector('span.middle_fade');
                                const toEmail = toSpan ? toSpan.textContent.trim() : 'N/A';

                                // Objet - dans la ligne de détails s_details_XXXX
                                const detailsRow = doc.querySelector(`#s_details_${emailId}`);
                                let subject = 'N/A';
                                if (detailsRow) {
                                    const subjectTd = detailsRow.querySelector('td[data-sent_email_id]');
                                    if (subjectTd) {
                                        subject = subjectTd.textContent.trim();
                                    }
                                }

                                // Pièce jointe
                                const hasAttachment = !!row.querySelector('.fa-paperclip');

                                // Récupérer le corps de l'email
                                updateLoader(`Lecture email ${emailCount}...`);
                                const body = await this.getEmailBody(emailId);
                                await Utils.delay(CONFIG.DELAY_EMAIL_BODY);

                                results.push({
                                    id: emailId,
                                    date: dateText,
                                    time: emailTime,
                                    toEmail: toEmail,
                                    subject: subject,
                                    body: body,
                                    hasAttachment: hasAttachment
                                });

                                Utils.log(`Email collecté: ${emailId} -> ${toEmail} | ${subject}`);
                            }
                            // Si l'email est AVANT la date du rapport → on a dépassé, arrêter
                            else if (emailDateObj < reportDateObj) {
                                Utils.log(`Email ${emailDate} antérieur à ${reportDate}, arrêt`);
                                passedReportDate = true;
                                break;
                            }
                        }
                    }

                    // Vérifier pagination - continuer tant qu'on n'a pas dépassé la date du rapport
                    const nextPageLink = doc.querySelector(`a[href*="sent_email_page=${currentPage + 1}"]`);
                    if (!nextPageLink || emailRows.length === 0 || passedReportDate) {
                        hasMorePages = false;
                    } else {
                        currentPage++;
                        await Utils.delay(CONFIG.DELAY_BETWEEN_REQUESTS);
                    }
                }

                Utils.log(`Total: ${results.length} emails envoyés pour le ${reportDate}`);
            } catch (error) {
                Utils.log('Erreur collecte emails envoyés:', error);
            }

            return results;
        },

        async getEmailBody(emailId) {
            try {
                // Le contenu de l'email est dans une iframe: sent_emails_frame.php
                const url = `https://courtage.modulr.fr/fr/scripts/sent_emails/sent_emails_frame.php?sent_email_id=${emailId}`;
                Utils.log(`Récupération corps email ${emailId} depuis iframe: ${url}`);

                const html = await Utils.fetchPage(url);
                Utils.log(`HTML iframe reçu (500 chars): ${html.substring(0, 500)}`);

                // Le contenu de l'iframe est directement le corps de l'email
                // Nettoyer le HTML pour extraire le texte
                const doc = Utils.parseHTML(html);

                // Chercher le body ou le contenu principal
                const body = doc.body || doc.querySelector('body');
                if (body) {
                    // Nettoyer le texte
                    let text = body.innerHTML || '';
                    text = text
                        .replace(/<br\s*\/?>/gi, '\n')
                        .replace(/<\/p>/gi, '\n')
                        .replace(/<\/div>/gi, '\n')
                        .replace(/<[^>]+>/g, '')
                        .replace(/&nbsp;/g, ' ')
                        .replace(/&amp;/g, '&')
                        .replace(/&lt;/g, '<')
                        .replace(/&gt;/g, '>')
                        .replace(/&quot;/g, '"')
                        .replace(/&#39;/g, "'")
                        .replace(/\n{3,}/g, '\n\n')
                        .trim();

                    if (text && text.length > 10) {
                        Utils.log(`Corps email ${emailId} trouvé (${text.length} chars): ${text.substring(0, 100)}...`);
                        return text;
                    }
                }

                // Fallback: extraire tout le texte du HTML avec regex
                let fallbackText = html
                    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
                    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
                    .replace(/<br\s*\/?>/gi, '\n')
                    .replace(/<\/p>/gi, '\n')
                    .replace(/<\/div>/gi, '\n')
                    .replace(/<[^>]+>/g, '')
                    .replace(/&nbsp;/g, ' ')
                    .replace(/&amp;/g, '&')
                    .replace(/\n{3,}/g, '\n\n')
                    .trim();

                if (fallbackText && fallbackText.length > 10) {
                    Utils.log(`Corps email ${emailId} trouvé via fallback (${fallbackText.length} chars)`);
                    return fallbackText;
                }

                Utils.log(`Aucun corps trouvé pour email ${emailId}`);
                return '';
            } catch (error) {
                Utils.log('Erreur lecture corps email:', error);
                return '';
            }
        }
    };

  // ============================================
    // COLLECTEUR D'EMAILS AFFECTÉS (v4.8.2 - POST+GET)
    // ============================================
    const EmailsAffectedCollector = {
        async collect(connectedUser, updateLoader) {
            console.log('%c=== COLLECTE EMAILS AFFECTÉS (v4.8.2) ===', 'background: #4CAF50; color: white; padding: 5px;');
            console.log('Utilisateur connecté:', connectedUser);
            const results = [];
            const reportDate = Utils.getTodayDate();
            console.log('Date du rapport:', reportDate);

            const MAX_PAGES = 20;

            try {
                // 1. POST pour activer le filtre
                updateLoader('Activation filtre emails...');
                await fetch('https://courtage.modulr.fr/fr/scripts/emails/emails_list.php', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/x-www-form-urlencoded'},
                    body: 'action=filter&emails_filters%5Bshow_associated_emails%5D=1&mailbox_id=all',
                    credentials: 'include'
                });
                console.log('Filtre activé');

                // 2. GET toutes les pages en parallèle
                updateLoader(`Emails affectés - Pages 1-${MAX_PAGES}...`);

                const promises = [];
                for (let page = 1; page <= MAX_PAGES; page++) {
                    promises.push(
                        fetch(`https://courtage.modulr.fr/fr/scripts/emails/emails_list.php?email_page=${page}`, {
                            credentials: 'include'
                        })
                        .then(r => r.text())
                        .then(html => ({ page, html }))
                        .catch(err => ({ page, html: '', error: err }))
                    );
                }

                const responses = await Promise.all(promises);

                // 3. Traiter les résultats (triés par page)
                responses.sort((a, b) => a.page - b.page);

                for (const { page, html, error } of responses) {
                    if (error || !html) continue;

                    const doc = Utils.parseHTML(html);
                    const emailRows = doc.querySelectorAll('tr[id^="e_main_"]');

                    for (const row of emailRows) {
                        const emailId = row.id.replace('e_main_', '');

                        let affectedTo = '';
                        let affectedDate = '';
                        let affectedBy = '';

                        const hiddenSpans = row.querySelectorAll('span.hidden');
                        for (const span of hiddenSpans) {
                            const txt = span.textContent.trim();
                            const match = txt.match(/Affecté\s+à\s+(.+?)\s+par\s+(.+?)\s+le\s+(\d{2}\/\d{2}\/\d{4})/i);
                            if (match) {
                                affectedTo = match[1].trim();
                                affectedBy = match[2].trim();
                                affectedDate = match[3];
                                break;
                            }
                        }

                        if (!affectedBy) continue;
                        if (affectedDate !== reportDate) continue;

                        // Filtre par utilisateur
                        const userLower = connectedUser.toLowerCase().trim();
                        const byLower = affectedBy.toLowerCase().trim();

                        let isMatch = (byLower === userLower);
                        if (!isMatch) isMatch = byLower.includes(userLower) || userLower.includes(byLower);
                        if (!isMatch) {
    const byParts = byLower.split(/[\s,]+/).filter(p => p.length > 2);
    const userParts = userLower.split(/[\s,]+/).filter(p => p.length > 2);

    // Exiger que le PRÉNOM corresponde (pas juste le nom de famille)
    if (byParts.length > 0 && userParts.length > 0) {
        const byFirstName = byParts[0];
        const userFirstName = userParts[0];
        if (byFirstName === userFirstName ||
            byFirstName.includes(userFirstName) ||
            userFirstName.includes(byFirstName)) {
            isMatch = true;
        }
    }
}

                        if (!isMatch) continue;

                        const dateTimeSpan = row.querySelector('span[id^="e_datetime_"]');
                        let emailTime = '';
                        if (dateTimeSpan) {
                            const timeMatch = dateTimeSpan.textContent.match(/(\d{1,2}:\d{2})/);
                            if (timeMatch) emailTime = timeMatch[1];
                        }

                        const fromSpan = row.querySelector('span[id^="e_from_"]');
                        const fromText = fromSpan ? fromSpan.textContent.trim() : 'N/A';

                        let fromEmail = '';
                        const emailInput = row.querySelector('input.association_email_email');
                        if (emailInput) fromEmail = emailInput.value;

                        let subject = 'N/A';
                        const subjectInput = row.querySelector('input.association_email_subject');
                        if (subjectInput && subjectInput.value) subject = subjectInput.value;

                        if (!results.find(r => r.id === emailId)) {
                            results.push({
                                id: emailId,
                                date: affectedDate,
                                time: emailTime,
                                from: fromText,
                                fromEmail: fromEmail,
                                subject: subject,
                                affectedTo: affectedTo,
                                hasAttachment: !!row.querySelector('.fa-paperclip')
                            });
                            console.log(`%c  ✓ Page ${page}: ${emailId} → ${affectedTo}`, 'color: #4CAF50');
                        }
                    }
                }

                console.log(`%c=== RÉSULTAT: ${results.length} emails ===`, 'background: #4CAF50; color: white; padding: 5px;');
            } catch (error) {
                console.error('Erreur collecte emails affectés:', error);
            }

            return results;
        }
    };

    // ============================================
    // COLLECTEUR NOMBRE D'EMAILS EN ATTENTE
    // ============================================
    // Récupère le nombre d'emails assignés (en attente) pour l'utilisateur
    // depuis la liste des utilisateurs dans le menu d'affectation
    const PendingEmailsCollector = {
        async collect(connectedUser, updateLoader) {
            Utils.log('=== COLLECTE EMAILS EN ATTENTE ===');
            Utils.log('Utilisateur recherché:', connectedUser);

            try {
                updateLoader('Récupération emails en attente...');

                // Charger la page des emails SANS filtre (pour voir les non traités)
                const url = 'https://courtage.modulr.fr/fr/scripts/emails/emails_list.php?email_page=1';
                const html = await Utils.fetchPage(url);
                const doc = Utils.parseHTML(html);

                // Chercher dans les liens d'affectation le pattern "NOM (XX)"
                // Le lien peut être /emails/assign/ ou /intranet/emails/assign/
                const assignLinks = doc.querySelectorAll('a[href*="emails/assign"]');
                Utils.log(`${assignLinks.length} liens d'affectation trouvés`);

                const userLower = connectedUser.toLowerCase().trim();
                const userParts = userLower.split(/[\s,]+/).filter(p => p.length > 2);
                let pendingCount = 0;
                let foundUser = false;

                for (const link of assignLinks) {
                    // Récupérer le texte en nettoyant les espaces et caractères spéciaux
                    let text = link.textContent.trim();
                    // Supprimer les espaces multiples et &nbsp;
                    text = text.replace(/\s+/g, ' ').trim();

                    Utils.log(`  Lien brut: "${text}"`);

                    // Pattern: "Nom Prénom (XX)" - chercher le nombre entre parenthèses
                    // Le nom peut contenir des espaces, donc on cherche tout avant les parenthèses
                    const match = text.match(/^(.+?)\s*\((\d+)\)\s*$/);
                    if (match) {
                        const userName = match[1].trim();
                        const count = parseInt(match[2]);
                        const nameLower = userName.toLowerCase().trim();

                        Utils.log(`    Parsé: nom="${userName}", count=${count}`);

                        // Vérifier si c'est l'utilisateur connecté
                        let isMatch = false;

                        // Match exact
                        if (nameLower === userLower) {
                            isMatch = true;
                            Utils.log(`    -> Match exact`);
                        }
                        // Match inclusion
                        if (!isMatch && (nameLower.includes(userLower) || userLower.includes(nameLower))) {
                            isMatch = true;
                            Utils.log(`    -> Match inclusion`);
                        }
                        // Match par parties du nom (prénom OU nom)
                        if (!isMatch) {
                            const nameParts = nameLower.split(/[\s,]+/).filter(p => p.length > 2);
                            let matchedParts = 0;
                            for (const np of nameParts) {
                                for (const up of userParts) {
                                    if (np === up) {
                                        matchedParts++;
                                        break;
                                    }
                                }
                            }
                            // Si au moins une partie du nom correspond
                            if (matchedParts > 0) {
                                isMatch = true;
                                Utils.log(`    -> Match par parties (${matchedParts} correspondances)`);
                            }
                        }

                        if (isMatch) {
                            pendingCount = count;
                            foundUser = true;
                            Utils.log(`✓ MATCH TROUVÉ: "${userName}" = ${pendingCount} emails en attente`);
                            break; // Prendre le premier match
                        }
                    } else {
                        Utils.log(`    -> Pas de pattern (XX) trouvé`);
                    }
                }

                if (!foundUser) {
                    Utils.log(`✗ Aucun match trouvé pour "${connectedUser}" parmi les ${assignLinks.length} liens`);
                }

                Utils.log(`=== RÉSULTAT: ${pendingCount} emails en attente ===`);
                return pendingCount;

            } catch (error) {
                Utils.log('Erreur collecte emails en attente:', error);
                return 0;
            }
        }
    };

    // ============================================
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

    // ============================================
    // COLLECTEUR DE TÂCHES TERMINÉES
    // ============================================
    const TasksCompletedCollector = {
        async collect(userId, connectedUser, updateLoader) {
            Utils.log('Collecte des tâches terminées par', connectedUser);
            const results = [];
            const today = Utils.getTodayDate();

            try {
                updateLoader('Tâches terminées...');

                const baseUrl = 'https://courtage.modulr.fr/fr/scripts/Tasks/TasksList.php';
                const params = new URLSearchParams({
                    'tasks_filters[task_recipient]': userId.taskValue,
                    'tasks_filters[task_status]': 'finished'
                });

                const url = `${baseUrl}?${params.toString()}#entity_menu_task=0`;
                Utils.log('URL tâches terminées:', url);

                const html = await Utils.fetchPage(url);
                const doc = Utils.parseHTML(html);

                // Récupérer toutes les lignes du tableau
                const allRows = Array.from(doc.querySelectorAll('tr'));
                const taskRows = allRows.filter(row => row.id && row.id.startsWith('task:'));
                Utils.log(`${taskRows.length} tâches trouvées`);

                let taskCount = 0;

                for (let i = 0; i < taskRows.length; i++) {
                    const row = taskRows[i];
                    const dateCell = row.querySelector('td.align_center');
                    const dateSpan = dateCell ? dateCell.querySelector('span:last-child') : null;
                    const completedDate = dateSpan ? dateSpan.textContent.trim() : '';

                    Utils.log(`Tâche date: ${completedDate}, today: ${today}`);

                    // Vérifier si terminée aujourd'hui
                    if (completedDate.includes(today)) {
                        taskCount++;
                        const taskId = row.id.replace('task:', '');

                        updateLoader(`Lecture tâche ${taskCount}...`);

                        const titleSpan = row.querySelector('span.font_size_higher');
                        const clientLink = row.querySelector('a[href*="clients_card"]');

                        // Chercher le contenu dans la ligne suivante
                        // La ligne de contenu a la classe task_ended_background_color ou task_bg_color
                        // et contient td[colspan] avec un <p>
                        let content = '';

                        // Méthode 1: Chercher la ligne suivante dans le DOM
                        const rowIndex = allRows.indexOf(row);
                        if (rowIndex >= 0 && rowIndex < allRows.length - 1) {
                            const nextRow = allRows[rowIndex + 1];
                            Utils.log(`Ligne suivante classe: ${nextRow.className}`);

                            // Vérifier si c'est une ligne de contenu (pas une ligne task:)
                            if (!nextRow.id || !nextRow.id.startsWith('task:')) {
                                const contentCell = nextRow.querySelector('td[colspan] p');
                                if (contentCell) {
                                    content = Utils.cleanText(contentCell.innerHTML);
                                    Utils.log(`Contenu trouvé (${content.length} chars): ${content.substring(0, 80)}...`);
                                }
                            }
                        }

                        // Méthode 2: Si pas trouvé, chercher avec regex dans le HTML brut
                        if (!content) {
                            const taskIdPattern = new RegExp(`id="task:${taskId}"[\\s\\S]*?<tr[^>]*>\\s*<td[^>]*colspan[^>]*>\\s*<p[^>]*>([\\s\\S]*?)<\\/p>`, 'i');
                            const match = html.match(taskIdPattern);
                            if (match) {
                                content = Utils.cleanText(match[1]);
                                Utils.log(`Contenu trouvé via regex (${content.length} chars)`);
                            }
                        }

                        // Méthode 3: Aller chercher sur la page de la tâche
                        if (!content) {
                            Utils.log(`Pas de contenu trouvé dans la liste, récupération page tâche ${taskId}`);
                            const taskDetails = await this.getTaskDetails(taskId);
                            content = taskDetails.content || '';
                            await Utils.delay(CONFIG.DELAY_BETWEEN_REQUESTS);
                        }

                        // Parser les infos de création et dernière modification
                        let createdBy = 'N/A', createdDate = 'N/A';
                        let closedTime = ''; // Heure de clôture = dernière modification
                        let closedBy = '';
                        const hiddenDiv = row.querySelector('.hidden');
                        if (hiddenDiv) {
                            const text = hiddenDiv.innerHTML;
                            // Extraction création
                            const creationMatch = text.match(/Création<\/p>\s*<p[^>]*>([^<]+)/);
                            if (creationMatch) {
                                const parts = creationMatch[1].trim().match(/(.+) (\d{2}\/\d{2}\/\d{4})/);
                                if (parts) {
                                    createdBy = parts[1].trim();
                                    createdDate = parts[2];
                                }
                            }
                            // Extraction dernière modification (= heure de clôture)
                            const modifMatch = text.match(/Derni[èe]re modification<\/p>\s*<p[^>]*>([^<]+)/i);
                            if (modifMatch) {
                                // Format: "NOM PRENOM DD/MM/YYYY HH:MM:SS"
                                const modifParts = modifMatch[1].trim().match(/(.+?)\s+(\d{2}\/\d{2}\/\d{4})\s+(\d{2}:\d{2})/);
                                if (modifParts) {
                                    closedBy = modifParts[1].trim();
                                    closedTime = modifParts[3]; // HH:MM
                                    Utils.log(`Tâche ${taskId}: Clôturée à ${closedTime} par ${closedBy}`);
                                }
                            }
                        }

                        results.push({
                            id: taskId,
                            title: titleSpan ? titleSpan.textContent.trim() : 'N/A',
                            content: content,
                            client: clientLink ? clientLink.textContent.trim() : 'Non associé',
                            clientId: clientLink ? (clientLink.href.match(/id=(\d+)/) || [])[1] : null,
                            assignedTo: connectedUser,
                            completedDate: completedDate,
                            time: closedTime, // Heure de clôture pour la vue chronologique
                            closedTime: closedTime,
                            closedBy: closedBy,
                            createdBy,
                            createdDate,
                            isPriority: !!row.querySelector('.fa-exclamation'),
                            hasBookmark: !!row.querySelector('.fa-bookmark')
                        });

                        Utils.log(`Tâche collectée: ${taskId}`);
                    }
                }

                Utils.log(`Total: ${results.length} tâches terminées`);
            } catch (error) {
                Utils.log('Erreur collecte tâches terminées:', error);
            }

            return results;
        },

        async getTaskDetails(taskId) {
            try {
                // L'URL de la popup de tâche
                const url = `https://courtage.modulr.fr/fr/scripts/Tasks/TasksCard.php?task_id=${taskId}`;
                Utils.log(`Récupération détails tâche ${taskId}: ${url}`);

                const html = await Utils.fetchPage(url);
                Utils.log(`HTML tâche reçu (300 chars): ${html.substring(0, 300)}`);

                let content = '';

                // Méthode 1: Chercher dans td[colspan] p (structure de la popup)
                // <tr><td colspan="4"><p class="medium_padding_left medium_padding_right">CONTENU</p></td></tr>
                const regexContent = /<td\s+colspan[^>]*>\s*<p[^>]*>([\s\S]*?)<\/p>\s*<\/td>/i;
                const match = html.match(regexContent);
                if (match) {
                    content = Utils.cleanText(match[1]);
                    Utils.log(`Contenu tâche trouvé via regex (${content.length} chars): ${content.substring(0, 80)}...`);
                }

                // Méthode 2: Parser le DOM
                if (!content) {
                    const doc = Utils.parseHTML(html);

                    // Chercher td[colspan] p
                    const contentCell = doc.querySelector('td[colspan] p');
                    if (contentCell) {
                        content = contentCell.innerHTML
                            .replace(/<br\s*\/?>/gi, '\n')
                            .replace(/<[^>]+>/g, '')
                            .trim();
                        Utils.log(`Contenu tâche trouvé via DOM: ${content.substring(0, 80)}...`);
                    }

                    // Fallback: textarea
                    if (!content) {
                        const textarea = doc.querySelector('textarea');
                        if (textarea && textarea.value) {
                            content = textarea.value.trim();
                        }
                    }
                }

                return { content: content };
            } catch (error) {
                Utils.log('Erreur lecture tâche:', error);
                return { content: '' };
            }
        }
    };

    // ============================================
    // COLLECTEUR DE TÂCHES EN RETARD
    // ============================================
    const TasksOverdueCollector = {
        async collect(userId, connectedUser, updateLoader) {
            Utils.log('Collecte des tâches en retard pour', connectedUser);
            const results = [];

            try {
                updateLoader('Tâches en retard...');

                // URL des tâches non terminées pour l'utilisateur
                const baseUrl = 'https://courtage.modulr.fr/fr/scripts/Tasks/TasksList.php';
                const params = new URLSearchParams({
                    'tasks_filters[task_recipient]': userId.taskValue,
                    'tasks_filters[task_status]': '' // Vide = toutes les tâches non terminées
                });

                const url = `${baseUrl}?${params.toString()}`;
                Utils.log('URL tâches en retard:', url);

                const html = await Utils.fetchPage(url);
                const doc = Utils.parseHTML(html);

                // Récupérer toutes les lignes pour pouvoir naviguer
                const allRows = Array.from(doc.querySelectorAll('tr'));
                const taskRows = allRows.filter(row => row.id && row.id.startsWith('task:'));
                Utils.log(`${taskRows.length} tâches trouvées au total`);

                let taskCount = 0;

                for (const row of taskRows) {
                    // Vérifier si la tâche est en retard
                    const isLate = row.classList.contains('task_late_background_color') ||
                                   row.querySelector('.task_late_icon') ||
                                   row.querySelector('.task_late_divider') ||
                                   row.querySelector('.fa-exclamation-triangle') ||
                                   row.querySelector('[class*="late"]') ||
                                   row.querySelector('[class*="overdue"]') ||
                                   row.style.backgroundColor?.includes('red') ||
                                   row.style.backgroundColor?.includes('ffcdd2');

                    // Alternative : vérifier la date d'échéance
                    const dateCell = row.querySelector('td.align_center');
                    let dueDate = '';
                    let isDatePast = false;

                    if (dateCell) {
                        const dateSpan = dateCell.querySelector('span:last-child');
                        if (dateSpan) {
                            dueDate = dateSpan.textContent.trim();
                            const daysOverdue = this.calculateDaysOverdue(dueDate);
                            isDatePast = daysOverdue > 0;
                        }
                    }

                    Utils.log(`Tâche ${row.id}: isLate=${isLate}, isDatePast=${isDatePast}, dueDate=${dueDate}`);

                    if (isLate || isDatePast) {
                        taskCount++;
                        const taskId = row.id.replace('task:', '');

                        const titleSpan = row.querySelector('span.font_size_higher');
                        const clientLink = row.querySelector('a[href*="clients_card"]');

                        const daysOverdue = this.calculateDaysOverdue(dueDate);

                        // Chercher le contenu dans la ligne suivante (comme pour les tâches terminées)
                        let content = '';
                        const rowIndex = allRows.indexOf(row);
                        if (rowIndex >= 0 && rowIndex < allRows.length - 1) {
                            const nextRow = allRows[rowIndex + 1];
                            if (!nextRow.id || !nextRow.id.startsWith('task:')) {
                                const contentCell = nextRow.querySelector('td[colspan] p');
                                if (contentCell) {
                                    content = Utils.cleanText(contentCell.innerHTML);
                                    Utils.log(`Contenu tâche retard trouvé: ${content.substring(0, 50)}...`);
                                }
                            }
                        }

                        // Si pas trouvé et moins de 10 tâches, aller chercher sur la page
                        if (!content && taskCount <= 10) {
                            updateLoader(`Lecture tâche retard ${taskCount}...`);
                            const taskDetails = await TasksCompletedCollector.getTaskDetails(taskId);
                            content = taskDetails.content || '';
                            await Utils.delay(CONFIG.DELAY_BETWEEN_REQUESTS);
                        }

                        results.push({
                            id: taskId,
                            title: titleSpan ? titleSpan.textContent.trim() : 'N/A',
                            content: content,
                            client: clientLink ? clientLink.textContent.trim() : 'Non associé',
                            clientId: clientLink ? (clientLink.href.match(/id=(\d+)/) || [])[1] : null,
                            assignedTo: connectedUser,
                            dueDate: dueDate,
                            daysOverdue: daysOverdue,
                            isPriority: !!row.querySelector('.fa-exclamation')
                        });

                        Utils.log(`Tâche en retard collectée: ${taskId} - ${daysOverdue}j`);
                    }
                }

                // Trier par retard décroissant
                results.sort((a, b) => b.daysOverdue - a.daysOverdue);
                Utils.log(`${results.length} tâches en retard trouvées`);
            } catch (error) {
                Utils.log('Erreur collecte tâches en retard:', error);
            }

            return results;
        },

        calculateDaysOverdue(dateStr) {
            if (!dateStr || dateStr === 'N/A') return 0;

            // Parser la date (format DD/MM/YYYY ou DD/MM/YYYY à HH:MM)
            const cleanDate = dateStr.split(' à ')[0].split(' ')[0].trim();
            const parts = cleanDate.split('/');
            if (parts.length < 3) return 0;

            const day = parseInt(parts[0], 10);
            const month = parseInt(parts[1], 10) - 1;
            const year = parseInt(parts[2], 10);

            if (isNaN(day) || isNaN(month) || isNaN(year)) return 0;

            const dueDate = new Date(year, month, day);

            // Utiliser la date du rapport pour calculer le retard
            let reportDate;
            if (SELECTED_REPORT_DATE) {
                const rParts = SELECTED_REPORT_DATE.split('/');
                reportDate = new Date(parseInt(rParts[2]), parseInt(rParts[1]) - 1, parseInt(rParts[0]));
            } else {
                reportDate = new Date();
            }
            reportDate.setHours(0, 0, 0, 0);

            const diffTime = reportDate - dueDate;
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

            return diffDays > 0 ? diffDays : 0;
        }
    };

    // ============================================
    // COLLECTEUR DE JOURNALISATION (VULGARISÉ)
    // ============================================
    const LogsCollector = {
        // Collecter les logs pour une table spécifique AVEC PAGINATION
        async collectByTable(userId, today, tableName, tableLabel, updateLoader) {
            const results = [];
            let currentPage = 1;
            let hasMorePages = true;
            const pageFingerprints = new Set();

            try {
                while (hasMorePages) {
                    if (updateLoader) updateLoader(`${tableLabel} - Page ${currentPage}...`);

                    const baseUrl = 'https://courtage.modulr.fr/fr/scripts/UsersLogs/UsersLogsList.php';
                    const params = new URLSearchParams();
                    params.append('mcut', '');
                    params.append('filters[user_id]', userId.logValue);
                    params.append('filters[user_log_date]', today);
                    params.append('filters[table]', tableName);
                    params.append('page', currentPage); // PAGINATION

                    Utils.log(`POST ${tableLabel} page ${currentPage}: ${baseUrl}`);

                    const html = await Utils.fetchPagePost(baseUrl, params.toString());
                    const doc = Utils.parseHTML(html);

                    const noResult = html.includes('Aucun résultat') || html.includes('aucun résultat');
                    if (noResult && currentPage === 1) {
                        Utils.log(`Aucun résultat pour ${tableLabel}`);
                        break;
                    }

                    const tableEl = doc.querySelector('table.table_list');
                    if (!tableEl) {
                        Utils.log(`Pas de tableau pour ${tableLabel} page ${currentPage}`);
                        break;
                    }

                    const rows = tableEl.querySelectorAll('tr');
                    Utils.log(`Page ${currentPage}: ${rows.length} lignes pour ${tableLabel}`);

                    const fingerprint = Array.from(rows)
                        .filter(row => row.classList.contains('color_grey_3'))
                        .slice(0, 5)
                        .map(row => row.textContent.replace(/\s+/g, ' ').trim())
                        .join('|');
                    if (fingerprint && pageFingerprints.has(fingerprint)) {
                        Utils.log(`Page ${currentPage} identique à une page précédente : arrêt sécurisé.`);
                        break;
                    }
                    if (fingerprint) pageFingerprints.add(fingerprint);

                    // Compter les entrées ajoutées sur cette page
                    let entriesThisPage = 0;
                    let currentEntry = null;

                    for (const row of rows) {
                        const cells = row.querySelectorAll('td');
                        const isHeaderRow = row.classList.contains('color_grey_3') &&
                                           row.classList.contains('no_hover_background') &&
                                           cells.length === 6;

                        if (isHeaderRow) {
                            if (currentEntry) {
                                results.push(currentEntry);
                                entriesThisPage++;
                            }

                            const actionRaw = cells[1]?.textContent?.trim() || '';
                            if (actionRaw && (actionRaw.includes('Insertion') || actionRaw.includes('Mise à jour') || actionRaw.includes('Suppression'))) {
                                currentEntry = {
                                    user: cells[0]?.textContent?.trim() || 'N/A',
                                    actionRaw: actionRaw,
                                    action: this.translateAction(actionRaw),
                                    date: cells[2]?.textContent?.trim() || 'N/A',
                                    tableRaw: cells[3]?.textContent?.trim() || 'N/A',
                                    table: tableLabel,
                                    entityId: cells[4]?.textContent?.trim() || 'N/A',
                                    entityName: cells[5]?.textContent?.trim() || 'N/A',
                                    clientName: null,
                                    category: tableName,
                                    changes: []
                                };
                            } else {
                                currentEntry = null;
                            }
                        }
                        else if (currentEntry && cells.length >= 2) {
                            const redCell = row.querySelector('.background_light_red');
                            const greenCell = row.querySelector('.background_light_green');

                            if (redCell || greenCell) {
                                const fieldSpan = cells[0]?.querySelector('span.high_margin_left');
                                const fieldRaw = fieldSpan?.textContent?.trim() || cells[0]?.textContent?.trim() || '';
                                const systemFields = ['last_update', 'last_update_user_id', 'creation_date', 'creation_user_id',
                                                      'estimate_id', 'policy_id', 'claim_id', 'office_id', 'firm_id', 'bank_account_id'];

                                if (fieldRaw && !systemFields.includes(fieldRaw)) {
                                    currentEntry.changes.push({
                                        fieldRaw: fieldRaw,
                                        field: Utils.translateField(fieldRaw),
                                        oldValueRaw: redCell?.textContent?.trim() || '-',
                                        oldValue: Utils.translateValue(redCell?.textContent?.trim() || '-'),
                                        newValueRaw: greenCell?.textContent?.trim() || '-',
                                        newValue: Utils.translateValue(greenCell?.textContent?.trim() || '-')
                                    });
                                }
                            }
                        }
                    }

                    if (currentEntry) {
                        results.push(currentEntry);
                        entriesThisPage++;
                    }

                    Utils.log(`Page ${currentPage}: ${entriesThisPage} entrées ajoutées pour ${tableLabel}`);

                    // Vérifier s'il y a une page suivante
                    const nextPageLink = doc.querySelector('a[href*="page=' + (currentPage + 1) + '"]') ||
                                        doc.querySelector('.pagination a.next') ||
                                        doc.querySelector('a[title="Page suivante"]');

                    // Si moins de 50 entrées, probablement dernière page
                    if (entriesThisPage < 50 && !nextPageLink) {
                        hasMorePages = false;
                    } else if (entriesThisPage === 0) {
                        hasMorePages = false;
                    } else {
                        currentPage++;
                        await Utils.delay(CONFIG.DELAY_BETWEEN_REQUESTS);
                    }
                }

                Utils.log(`Total ${results.length} entrées pour ${tableLabel}`);

            } catch (error) {
                Utils.log(`Erreur collecte logs ${tableName}:`, error);
            }

            return results;
        },

        async collect(userId, connectedUser, updateLoader) {
            Utils.log('Collecte de la journalisation générale pour', connectedUser);
            const results = [];
            const today = Utils.getTodayDate();
            let currentPage = 1;
            let hasMorePages = true;
            const pageFingerprints = new Set();

            try {
                while (hasMorePages) {
                    updateLoader(`Journalisation générale - Page ${currentPage}...`);

                    const baseUrl = 'https://courtage.modulr.fr/fr/scripts/UsersLogs/UsersLogsList.php';
                    const params = new URLSearchParams();
                    params.append('filters[user_id]', userId.logValue);
                    params.append('filters[user_log_date]', today);
                    params.append('page', currentPage);

                    Utils.log(`POST journalisation générale page ${currentPage}`);

                    const html = await Utils.fetchPagePost(baseUrl, params.toString());
                    const doc = Utils.parseHTML(html);

                    const noResult = html.includes('Aucun résultat') || html.includes('aucun résultat');
                    if (noResult && currentPage === 1) {
                        Utils.log('Aucun résultat pour logs généraux');
                        break;
                    }

                    const tableEl = doc.querySelector('table.table_list');
                    if (!tableEl) {
                        Utils.log('Pas de tableau table_list pour logs généraux');
                        break;
                    }

                    const rows = tableEl.querySelectorAll('tr');
                    Utils.log(`Page ${currentPage}: ${rows.length} lignes pour logs généraux`);

                    const fingerprint = Array.from(rows)
                        .filter(row => row.classList.contains('color_grey_3'))
                        .slice(0, 5)
                        .map(row => row.textContent.replace(/\s+/g, ' ').trim())
                        .join('|');
                    if (fingerprint && pageFingerprints.has(fingerprint)) {
                        Utils.log(`Page ${currentPage} répétée par Modulr : arrêt sécurisé.`);
                        break;
                    }
                    if (fingerprint) pageFingerprints.add(fingerprint);

                    let entriesThisPage = 0;
                    let currentEntry = null;

                    for (const row of rows) {
                        const cells = row.querySelectorAll('td');
                        const isHeaderRow = row.classList.contains('color_grey_3') &&
                                           row.classList.contains('no_hover_background') &&
                                           cells.length === 6;

                        if (isHeaderRow) {
                            if (currentEntry) {
                                results.push(currentEntry);
                                entriesThisPage++;
                            }

                            const actionRaw = cells[1]?.textContent?.trim() || '';
                            const tableRaw = cells[3]?.textContent?.trim() || '';

                            if (actionRaw && (actionRaw.includes('Insertion') || actionRaw.includes('Mise à jour') || actionRaw.includes('Suppression'))) {
                                currentEntry = {
                                    user: cells[0]?.textContent?.trim() || 'N/A',
                                    actionRaw: actionRaw,
                                    action: this.translateAction(actionRaw),
                                    date: cells[2]?.textContent?.trim() || 'N/A',
                                    tableRaw: tableRaw,
                                    table: Utils.translateTable(tableRaw),
                                    entityId: cells[4]?.textContent?.trim() || 'N/A',
                                    entityName: cells[5]?.textContent?.trim() || 'N/A',
                                    clientName: null,
                                    category: 'general',
                                    changes: []
                                };
                            } else {
                                currentEntry = null;
                            }
                        }
                        else if (currentEntry && cells.length >= 2) {
                            const redCell = row.querySelector('.background_light_red');
                            const greenCell = row.querySelector('.background_light_green');

                            if (redCell || greenCell) {
                                const fieldSpan = cells[0]?.querySelector('span.high_margin_left');
                                const fieldRaw = fieldSpan?.textContent?.trim() || cells[0]?.textContent?.trim() || '';

                                const systemFields = ['last_update', 'last_update_user_id', 'creation_date', 'creation_user_id',
                                                      'estimate_id', 'policy_id', 'claim_id', 'office_id', 'firm_id', 'bank_account_id'];
                                if (fieldRaw && !systemFields.includes(fieldRaw)) {
                                    currentEntry.changes.push({
                                        fieldRaw: fieldRaw,
                                        field: Utils.translateField(fieldRaw),
                                        oldValueRaw: redCell?.textContent?.trim() || '-',
                                        oldValue: Utils.translateValue(redCell?.textContent?.trim() || '-'),
                                        newValueRaw: greenCell?.textContent?.trim() || '-',
                                        newValue: Utils.translateValue(greenCell?.textContent?.trim() || '-')
                                    });
                                }
                            }
                        }
                    }

                    if (currentEntry) {
                        results.push(currentEntry);
                        entriesThisPage++;
                    }

                    Utils.log(`Page ${currentPage}: ${entriesThisPage} entrées générales ajoutées`);

                    // Pagination
                    if (entriesThisPage < 50 || entriesThisPage === 0) {
                        hasMorePages = false;
                    } else {
                        currentPage++;
                        await Utils.delay(CONFIG.DELAY_BETWEEN_REQUESTS);
                    }
                }

                Utils.log(`Total ${results.length} actions générales trouvées`);
            } catch (error) {
                Utils.log('Erreur collecte journalisation générale:', error);
            }

            return this.keepUsefulUniqueLogs(results);
        },

        // La journalisation brute contient aussi les emails, tâches, devis, contrats
        // et sinistres déjà affichés dans leurs rubriques. On ne garde ici que les
        // événements complémentaires, en supprimant les doublons et le bruit système.
        keepUsefulUniqueLogs(results) {
            const alreadyReported = /(^|[^a-z])(tasks?|tâches?|emails?|sent_emails|estimates?|devis|polic(?:y|ies)|contrats?|claims?|sinistres?)([^a-z]|$)/i;
            const technicalFields = new Set([
                'last_update', 'last_update_user_id', 'creation_date', 'creation_user_id',
                'office_id', 'firm_id', 'blob_id', 'uid', 'eml_message_id',
                'tracking_data_algorithm', 'email_id', 'task_id', 'estimate_id',
                'policy_id', 'claim_id', 'bank_account_id'
            ]);
            const seen = new Set();

            return (results || []).filter(log => {
                const table = `${log.tableRaw || ''} ${log.table || ''}`.toLowerCase();
                if (alreadyReported.test(table)) return false;

                log.changes = (log.changes || []).filter(change => !technicalFields.has(change.fieldRaw));
                const isMeaningfulCreation = /insertion|création|insert/i.test(log.actionRaw || log.action || '');
                if (!isMeaningfulCreation && log.changes.length === 0) return false;

                const key = [log.actionRaw, log.tableRaw, log.entityId, log.date,
                    log.changes.map(c => `${c.fieldRaw}:${c.oldValueRaw}:${c.newValueRaw}`).join('|')].join('::');
                if (seen.has(key)) return false;
                seen.add(key);
                return true;
            });
        },

        // Collecter les devis
        async collectEstimates(userId, connectedUser, updateLoader) {
            Utils.log('Collecte des devis pour', connectedUser);
            const today = Utils.getTodayDate();

            const estimates = await this.collectByTable(userId, today, 'estimates', 'Devis', updateLoader);

            Utils.log(`${estimates.length} actions sur devis trouvées`);
            return estimates;
        },

        // Collecter les contrats
        async collectPolicies(userId, connectedUser, updateLoader) {
            Utils.log('Collecte des contrats pour', connectedUser);
            const today = Utils.getTodayDate();

            const policies = await this.collectByTable(userId, today, 'policies', 'Contrats', updateLoader);

            Utils.log(`${policies.length} actions sur contrats trouvées`);
            return policies;
        },

        // Collecter les sinistres
        async collectClaims(userId, connectedUser, updateLoader) {
            Utils.log('Collecte des sinistres pour', connectedUser);
            const today = Utils.getTodayDate();

            const claims = await this.collectByTable(userId, today, 'claims', 'Sinistres', updateLoader);

            Utils.log(`${claims.length} actions sur sinistres trouvées`);
            return claims;
        },

        translateAction(action) {
            const translations = {
                'Insertion': '➕ Création',
                'Mise à jour': '✏️ Modification',
                'Suppression': '🗑️ Suppression',
                'Delete': '🗑️ Suppression',
                'Update': '✏️ Modification',
                'Insert': '➕ Création'
            };
            return translations[action] || action;
        }
    };

    // ============================================
    // RÉSOLVEUR DE CLIENTS (Correspondance Email <-> N° Client <-> Nom)
    // ============================================
    const ClientResolver = {
        // Index des clients : clé (email/id/nom) -> {id, name, email}
        clientIndex: new Map(),

        async resolve(data, updateLoader) {
            Utils.log('=== DÉBUT RÉSOLUTION DES CLIENTS ===');

            // 1. Extraire tous les identifiants à rechercher avec contexte
            const searchItems = [];

            // Depuis les emails envoyés - on a l'email du destinataire
            data.emailsSent.forEach(e => {
                if (e.toEmail && !this.isInternalEmail(e.toEmail)) {
                    searchItems.push({
                        type: 'email',
                        value: e.toEmail.toLowerCase(),
                        context: e.subject || '',
                        source: e
                    });
                }
            });

            // Depuis les emails affectés - on a l'email de l'expéditeur ET le nom du client affecté
            data.emailsAffected.forEach(e => {
                // Le client auquel c'est affecté est dans affectedTo
                if (e.affectedTo && e.affectedTo !== 'N/A') {
                    searchItems.push({
                        type: 'name',
                        value: e.affectedTo,
                        context: e.subject || '',
                        source: e
                    });
                }
                // L'expéditeur peut aussi être un client
                if (e.fromEmail && !this.isInternalEmail(e.fromEmail)) {
                    searchItems.push({
                        type: 'email',
                        value: e.fromEmail.toLowerCase(),
                        context: e.subject || '',
                        source: e
                    });
                }
            });

            // Depuis les tâches - on a le nom du client et parfois l'ID
            data.tasksCompleted.forEach(t => {
                if (t.clientId) {
                    searchItems.push({ type: 'id', value: t.clientId, source: t });
                } else if (t.client && t.client !== 'Non associé') {
                    searchItems.push({ type: 'name', value: t.client, source: t });
                }
            });
            data.tasksOverdue.forEach(t => {
                if (t.clientId) {
                    searchItems.push({ type: 'id', value: t.clientId, source: t });
                } else if (t.client && t.client !== 'Non associé') {
                    searchItems.push({ type: 'name', value: t.client, source: t });
                }
            });

            // Depuis les logs - extraire le N° depuis entityName (format "n° XXXX du DD/MM/YYYY")
            [...data.estimates, ...data.policies, ...data.claims, ...data.logs].forEach(log => {
                if (log.entityId) {
                    searchItems.push({ type: 'id', value: log.entityId, source: log });
                } else if (log.entityName) {
                    // Essayer d'extraire un N° client
                    const numMatch = log.entityName.match(/n°\s*(\d+)/i);
                    if (numMatch) {
                        searchItems.push({ type: 'id', value: numMatch[1], source: log });
                    }
                }
            });

            // 2. Dédupliquer les recherches
            const uniqueSearches = new Map();
            searchItems.forEach(item => {
                const key = `${item.type}:${item.value}`;
                if (!uniqueSearches.has(key)) {
                    uniqueSearches.set(key, item);
                }
            });

            Utils.log(`${uniqueSearches.size} recherches uniques à effectuer`);

            // 3. Effectuer les recherches
            let searchCount = 0;
            const totalSearches = uniqueSearches.size;

            for (const [key, item] of uniqueSearches) {
                // Vérifier si déjà dans l'index
                if (this.clientIndex.has(item.value.toString().toLowerCase())) {
                    continue;
                }

                searchCount++;
                updateLoader(`Résolution client ${searchCount}/${totalSearches}...`);

                try {
                    let clientInfo = null;

                    if (item.type === 'id') {
                        // Recherche directe par ID - va sur la fiche client
                        clientInfo = await this.fetchClientById(item.value);
                    } else {
                        // Recherche globale par email ou nom
                        clientInfo = await this.searchGlobal(item.value, item.context);
                    }

                    if (clientInfo) {
                        // Indexer par ID, email et nom
                        this.clientIndex.set(clientInfo.id.toString(), clientInfo);
                        if (clientInfo.email) {
                            this.clientIndex.set(clientInfo.email.toLowerCase(), clientInfo);
                        }
                        if (clientInfo.name) {
                            this.clientIndex.set(clientInfo.name.toLowerCase(), clientInfo);
                        }
                        Utils.log(`✓ Client trouvé: ${clientInfo.name} (N° ${clientInfo.id}) - ${clientInfo.email || 'pas d\'email'}`);
                    }
                } catch (searchError) {
                    Utils.log(`Erreur recherche pour ${item.value}:`, searchError.message || searchError);
                    // Continuer avec les autres recherches
                }

                await Utils.delay(150); // Éviter de surcharger le serveur
            }

            Utils.log(`Index clients: ${this.clientIndex.size} entrées`);

            // 4. Enrichir les données
            this.enrichData(data);

            Utils.log('=== FIN RÉSOLUTION DES CLIENTS ===');
            return data;
        },

        // Vérifier si c'est un email interne (LTOA, etc.)
        isInternalEmail(email) {
            if (!email) return true;
            const internal = ['ltoa.fr', 'ltoaassurances.fr', 'modulr.fr'];
            return internal.some(domain => email.toLowerCase().includes(domain));
        },

        // Recherche globale Modulr
        async searchGlobal(query, context = '') {
            try {
                Utils.log(`Recherche globale: "${query}"`);

                // Utiliser la recherche globale Modulr
                const searchUrl = `https://courtage.modulr.fr/fr/scripts/user_global_search.php?global_search=${encodeURIComponent(query)}`;

                const response = await fetch(searchUrl, {
                    method: 'GET',
                    credentials: 'include',
                    headers: { 'Accept': 'text/html' }
                });

                if (!response.ok) {
                    Utils.log(`  Erreur HTTP ${response.status} pour recherche "${query}"`);
                    return null;
                }

                const html = await response.text();
                const doc = Utils.parseHTML(html);

                // Vérifier si on est directement sur une fiche client (titre contient le nom)
                const pageTitle = doc.querySelector('title')?.textContent || '';
                if (pageTitle.includes(' - Modulr') && !pageTitle.includes('Recherche')) {
                    // On est sur une fiche client directe
                    return this.parseClientCard(doc, html);
                }

                // Sinon, on a une liste de résultats - chercher dans le tableau
                const clientRows = doc.querySelectorAll('tr[id^="global_search_goto_client_card_"]');

                if (clientRows.length === 0) {
                    Utils.log(`  Aucun résultat pour "${query}"`);
                    return null;
                }

                if (clientRows.length === 1) {
                    // Un seul résultat - l'utiliser directement
                    return this.parseClientRow(clientRows[0]);
                }

                // Plusieurs résultats - essayer de départager avec le contexte
                Utils.log(`  ${clientRows.length} résultats, tentative de départage...`);

                // Chercher le meilleur match basé sur le contexte (objet du mail)
                let bestMatch = null;
                let bestScore = 0;

                for (const row of clientRows) {
                    const clientInfo = this.parseClientRow(row);
                    if (!clientInfo) continue;

                    // Calculer un score de correspondance
                    let score = 0;
                    const contextLower = context.toLowerCase();
                    const nameParts = clientInfo.name.toLowerCase().split(/[\s,]+/);

                    for (const part of nameParts) {
                        if (part.length > 2 && contextLower.includes(part)) {
                            score += 10;
                        }
                    }

                    // Si l'email correspond exactement à la recherche
                    if (clientInfo.email && clientInfo.email.toLowerCase() === query.toLowerCase()) {
                        score += 50;
                    }

                    if (score > bestScore) {
                        bestScore = score;
                        bestMatch = clientInfo;
                    }
                }

                // Si on a trouvé un bon match, l'utiliser
                if (bestMatch && bestScore > 0) {
                    Utils.log(`  Meilleur match: ${bestMatch.name} (score: ${bestScore})`);
                    return bestMatch;
                }

                // Sinon, prendre le premier résultat par défaut
                Utils.log(`  Pas de match contexte, utilisation du premier résultat`);
                return this.parseClientRow(clientRows[0]);

            } catch (error) {
                Utils.log(`Erreur recherche globale "${query}":`, error);
                return null;
            }
        },

        // Parser une ligne de résultat de recherche
        parseClientRow(row) {
            try {
                // ID du client depuis l'id de la ligne: global_search_goto_client_card_3350_0
                const rowId = row.id || '';
                const idMatch = rowId.match(/client_card_(\d+)/);
                if (!idMatch) return null;

                const clientId = idMatch[1];

                // Nom - dans la 3ème colonne
                const nameCell = row.querySelector('td:nth-child(3) p');
                const clientName = nameCell ? nameCell.textContent.trim() : '';

                // Email - dans le tooltip (span.hidden)
                let clientEmail = '';
                const hiddenContent = row.querySelector('span.hidden');
                if (hiddenContent) {
                    const emailLink = hiddenContent.querySelector('a[href*="documents_send.php"]');
                    if (emailLink) {
                        clientEmail = emailLink.getAttribute('title') || emailLink.textContent.trim();
                    }
                }

                if (!clientId || !clientName) return null;

                return {
                    id: clientId,
                    name: clientName,
                    email: clientEmail || null
                };
            } catch (error) {
                Utils.log('Erreur parsing row:', error);
                return null;
            }
        },

        // Parser une fiche client complète
        parseClientCard(doc, html) {
            try {
                // ID client - dans input hidden
                const clientIdInput = doc.querySelector('input[name="client_id_"], input#client_id_');
                const clientId = clientIdInput ? clientIdInput.value : null;

                if (!clientId) {
                    // Essayer depuis l'URL ou autre
                    const match = html.match(/client_id[=:](\d+)/i);
                    if (!match) return null;
                }

                // Nom - dans le titre de la page ou h1
                let clientName = '';
                const pageTitle = doc.querySelector('title')?.textContent || '';
                const titleMatch = pageTitle.match(/^([^-]+)/);
                if (titleMatch) {
                    clientName = titleMatch[1].trim();
                }

                // Ou dans le h1
                if (!clientName) {
                    const h1 = doc.querySelector('h1.page_title');
                    if (h1) {
                        clientName = h1.textContent.replace(/^\s*\S+\s*/, '').trim(); // Enlever l'icône
                    }
                }

                // Email - dans la vcard
                let clientEmail = '';
                const emailLink = doc.querySelector('.vcard a[href*="documents_send.php"]');
                if (emailLink) {
                    clientEmail = emailLink.getAttribute('title') || emailLink.textContent.trim();
                }

                // Ou dans un input email
                if (!clientEmail) {
                    const emailInput = doc.querySelector('input[type="email"], input[name*="email"]');
                    if (emailInput && emailInput.value) {
                        clientEmail = emailInput.value;
                    }
                }

                return {
                    id: clientId || clientIdInput?.value,
                    name: clientName,
                    email: clientEmail || null
                };
            } catch (error) {
                Utils.log('Erreur parsing fiche client:', error);
                return null;
            }
        },

        // Récupérer un client par son ID directement
        async fetchClientById(clientId) {
            try {
                Utils.log(`Fetch client par ID: ${clientId}`);
                const url = `https://courtage.modulr.fr/fr/scripts/clients/clients_card.php?id=${clientId}`;

                const response = await fetch(url, {
                    method: 'GET',
                    credentials: 'include',
                    headers: { 'Accept': 'text/html' }
                });

                if (!response.ok) {
                    Utils.log(`  Erreur HTTP ${response.status} pour client ${clientId}`);
                    return null;
                }

                const html = await response.text();
                const doc = Utils.parseHTML(html);

                return this.parseClientCard(doc, html);
            } catch (error) {
                Utils.log(`Erreur fetch client ${clientId}:`, error.message || error);
                return null; // Ne pas propager l'erreur
            }
        },

        // Enrichir les données avec les correspondances trouvées
        enrichData(data) {
            Utils.log('Enrichissement des données...');

            // Emails envoyés
            data.emailsSent.forEach(e => {
                if (e.toEmail) {
                    const clientInfo = this.getClientInfo(e.toEmail);
                    if (clientInfo) {
                        e.clientId = clientInfo.id;
                        e.clientName = clientInfo.name;
                        e.clientEmail = clientInfo.email;
                        e.clientResolved = true;
                    }
                }
            });

            // Emails affectés - utiliser le nom de l'affectation
            data.emailsAffected.forEach(e => {
                // D'abord essayer avec affectedTo (nom du client)
                if (e.affectedTo) {
                    const clientInfo = this.getClientInfo(e.affectedTo);
                    if (clientInfo) {
                        e.clientId = clientInfo.id;
                        e.clientName = clientInfo.name;
                        e.clientEmail = clientInfo.email;
                        e.clientResolved = true;
                        return;
                    }
                }
                // Sinon essayer avec l'email de l'expéditeur
                if (e.fromEmail) {
                    const clientInfo = this.getClientInfo(e.fromEmail);
                    if (clientInfo) {
                        e.clientId = clientInfo.id;
                        e.clientName = clientInfo.name;
                        e.clientEmail = clientInfo.email;
                        e.clientResolved = true;
                    }
                }
            });

            // Tâches
            [...data.tasksCompleted, ...data.tasksOverdue].forEach(t => {
                if (t.clientId) {
                    const clientInfo = this.getClientInfo(t.clientId);
                    if (clientInfo) {
                        t.clientName = clientInfo.name;
                        t.clientEmail = clientInfo.email;
                        t.clientResolved = true;
                    }
                } else if (t.client) {
                    const clientInfo = this.getClientInfo(t.client);
                    if (clientInfo) {
                        t.clientId = clientInfo.id;
                        t.clientName = clientInfo.name;
                        t.clientEmail = clientInfo.email;
                        t.clientResolved = true;
                    }
                }
            });

            // Logs (devis, contrats, sinistres)
            [...data.estimates, ...data.policies, ...data.claims, ...data.logs].forEach(log => {
                if (log.entityId) {
                    const clientInfo = this.getClientInfo(log.entityId);
                    if (clientInfo) {
                        log.clientId = clientInfo.id;
                        log.clientName = clientInfo.name;
                        log.clientEmail = clientInfo.email;
                        log.clientResolved = true;
                    }
                }
            });
        },

        // Obtenir les infos client depuis l'index
        getClientInfo(identifier) {
            if (!identifier) return null;
            const key = identifier.toString().toLowerCase().trim();
            return this.clientIndex.get(key) || null;
        },

        // Réinitialiser l'index
        reset() {
            this.clientIndex.clear();
        }
    };

    // ============================================
    // GÉNÉRATEUR DE RAPPORT (UI)
    // ============================================
    const ReportGenerator = {
        data: {
            emailsSent: [],
            emailsAffected: [],
            aircallCalls: [],
            tasksCompleted: [],
            tasksOverdue: [],
            logs: [],
            estimates: [],
            policies: [],
            claims: [],
            user: '',
            date: ''
        },

        generateHTML() {
            const { emailsSent, emailsAffected, pendingEmailsCount, aircallCalls, tasksCompleted, tasksOverdue, logs, estimates, policies, claims, user, date, notes, aircallStatus } = this.data;

            // Vérifier si c'est un rapport pour un jour passé
            const realToday = Utils.getRealTodayDate();
            const isPastDate = date !== realToday;
            const dateLabel = isPastDate ? `📅 ${date} <span style="background: #ff9800; color: white; padding: 2px 8px; border-radius: 3px; font-size: 12px; margin-left: 8px;">Rapport rétrospectif</span>` : date;

            // Compteurs Aircall
            const aircallInbound = (aircallCalls || []).filter(c => c.type === 'entrant').length;
            const aircallOutbound = (aircallCalls || []).filter(c => c.type === 'sortant').length;
            const aircallAnswered = (aircallCalls || []).filter(c => c.answered !== false && !c.missedReason).length;
            const aircallTalkSeconds = (aircallCalls || []).reduce((sum, c) => sum + (Number(c.durationSeconds) || 0), 0);
            const formatDuration = seconds => `${Math.floor(seconds / 3600)}h ${Math.floor((seconds % 3600) / 60)}min`;
            const uniqueClients = new Set([
                ...emailsSent, ...emailsAffected, ...tasksCompleted,
                ...estimates, ...policies, ...claims, ...logs
            ].map(item => item.clientId || item.clientName || item.entityId)
              .filter(value => value && value !== 'N/A' && value !== 'Non associé')).size;
            const remainingWork = (pendingEmailsCount || 0) + tasksOverdue.length;

            // Générer un ID unique pour les toggles
            const uid = Date.now();

            return `
                <div id="ltoa-report-modal" style="
                    position: fixed;
                    top: 0;
                    left: 0;
                    width: 100%;
                    height: 100%;
                    background: rgba(0,0,0,0.85);
                    z-index: 999999;
                    overflow-y: auto;
                    font-family: Arial, sans-serif;
                ">
                    <div id="ltoa-report-content" style="
                        max-width: 1200px;
                        margin: 20px auto;
                        background: white;
                        border-radius: 10px;
                        padding: 30px;
                        box-shadow: 0 10px 50px rgba(0,0,0,0.3);
                    ">
                        <!-- Header -->
                        <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 3px solid #c62828; padding-bottom: 20px; margin-bottom: 30px;">
                            <div>
                                <h1 style="color: #c62828; margin: 0; font-size: 24px;">📊 Rapport d'Activité Quotidien</h1>
                                <p style="color: #666; margin: 5px 0 0 0; font-size: 16px;">
                                    <strong>${user}</strong> - ${dateLabel}
                                </p>
                            </div>
                            <div>
                                <button id="ltoa-view-by-client" style="
                                    background: #1565c0;
                                    color: white;
                                    border: none;
                                    padding: 12px 20px;
                                    border-radius: 5px;
                                    cursor: pointer;
                                    font-size: 13px;
                                    margin-right: 8px;
                                    font-weight: bold;
                                ">👤 Vue par Client</button>
                                <button id="ltoa-export-html" style="
                                    background: #e65100;
                                    color: white;
                                    border: none;
                                    padding: 12px 20px;
                                    border-radius: 5px;
                                    cursor: pointer;
                                    font-size: 13px;
                                    margin-right: 8px;
                                    font-weight: bold;
                                ">🌐 Exporter HTML</button>
                                <button id="ltoa-view-chrono" style="
                                    background: #9c27b0;
                                    color: white;
                                    border: none;
                                    padding: 12px 20px;
                                    border-radius: 5px;
                                    cursor: pointer;
                                    font-size: 13px;
                                    margin-right: 8px;
                                    font-weight: bold;
                                ">🕐 Vue Chronologique</button>
                                <button id="ltoa-close-report" style="
                                    background: #666;
                                    color: white;
                                    border: none;
                                    padding: 12px 20px;
                                    border-radius: 5px;
                                    cursor: pointer;
                                    font-size: 13px;
                                    font-weight: bold;
                                ">✕ Fermer</button>
                            </div>
                        </div>

                        ${aircallStatus?.state === 'error' ? `
                        <div style="margin:-10px 0 25px;padding:12px 16px;border-radius:8px;background:#ffebee;color:#b71c1c;font-size:12px;">
                            ⚠️ ${Utils.escapeHtml(aircallStatus.message || 'Échec de la collecte Aircall')}
                        </div>` : ''}

                        <!-- Alerte emails en attente -->
                        ${(pendingEmailsCount || 0) > 0 ? `
                        <div style="background: linear-gradient(135deg, #ffcccb 0%, #ff6b6b 100%); padding: 15px 20px; border-radius: 10px; margin-bottom: 20px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 15px rgba(255, 107, 107, 0.3);">
                            <div style="display: flex; align-items: center; gap: 15px;">
                                <div style="font-size: 40px;">📬</div>
                                <div>
                                    <div style="font-size: 14px; color: #7f0000; font-weight: bold;">Emails assignés à ${Utils.escapeHtml(user)}</div>
                                </div>
                            </div>
                            <div style="font-size: 48px; font-weight: bold; color: #b71c1c;">${pendingEmailsCount || 0}</div>
                        </div>
                        ` : ''}

                        <!-- Résumé en cartes -->
                        <div style="display: grid; grid-template-columns: repeat(5, 1fr); gap: 12px; margin-bottom: 30px;">
                            <div style="background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%); padding: 15px; border-radius: 10px; text-align: center; box-shadow: 0 2px 10px rgba(25, 118, 210, 0.2);">
                                <div style="font-size: 28px; font-weight: bold; color: #1976d2;">${emailsSent.length}</div>
                                <div style="color: #1976d2; font-weight: bold; font-size: 12px;">📤 Emails Envoyés</div>
                            </div>
                            <div style="background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%); padding: 15px; border-radius: 10px; text-align: center; box-shadow: 0 2px 10px rgba(56, 142, 60, 0.2);">
                                <div style="font-size: 28px; font-weight: bold; color: #388e3c;">${emailsAffected.length}</div>
                                <div style="color: #388e3c; font-weight: bold; font-size: 12px;">📥 Emails Affectés</div>
                            </div>
                            <div style="background: linear-gradient(135deg, #fff8e1 0%, #ffecb3 100%); padding: 15px; border-radius: 10px; text-align: center; box-shadow: 0 2px 10px rgba(255, 160, 0, 0.2);">
                                <div style="font-size: 28px; font-weight: bold; color: #ff8f00;">${(aircallCalls || []).length}</div>
                                <div style="color: #ff8f00; font-weight: bold; font-size: 11px;">📞 Appels (${aircallInbound}↓ ${aircallOutbound}↑)</div>
                            </div>
                            <div style="background: linear-gradient(135deg, #fff3e0 0%, #ffe0b2 100%); padding: 15px; border-radius: 10px; text-align: center; box-shadow: 0 2px 10px rgba(245, 124, 0, 0.2);">
                                <div style="font-size: 28px; font-weight: bold; color: #f57c00;">${tasksCompleted.length}</div>
                                <div style="color: #f57c00; font-weight: bold; font-size: 12px;">✅ Tâches Terminées</div>
                            </div>
                            <div style="background: linear-gradient(135deg, #ffebee 0%, #ffcdd2 100%); padding: 15px; border-radius: 10px; text-align: center; box-shadow: 0 2px 10px rgba(211, 47, 47, 0.2);">
                                <div style="font-size: 28px; font-weight: bold; color: #d32f2f;">${tasksOverdue.length}</div>
                                <div style="color: #d32f2f; font-weight: bold; font-size: 12px;">⚠️ Tâches en Retard</div>
                            </div>
                        </div>
                        <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 30px;">
                            <div style="background: linear-gradient(135deg, #e0f7fa 0%, #b2ebf2 100%); padding: 15px; border-radius: 10px; text-align: center; box-shadow: 0 2px 10px rgba(0, 151, 167, 0.2);">
                                <div style="font-size: 28px; font-weight: bold; color: #0097a7;">${estimates.length}</div>
                                <div style="color: #0097a7; font-weight: bold; font-size: 12px;">📋 Devis</div>
                            </div>
                            <div style="background: linear-gradient(135deg, #e8eaf6 0%, #c5cae9 100%); padding: 15px; border-radius: 10px; text-align: center; box-shadow: 0 2px 10px rgba(63, 81, 181, 0.2);">
                                <div style="font-size: 28px; font-weight: bold; color: #3f51b5;">${policies.length}</div>
                                <div style="color: #3f51b5; font-weight: bold; font-size: 12px;">📄 Contrats</div>
                            </div>
                            <div style="background: linear-gradient(135deg, #fce4ec 0%, #f8bbd9 100%); padding: 15px; border-radius: 10px; text-align: center; box-shadow: 0 2px 10px rgba(194, 24, 91, 0.2);">
                                <div style="font-size: 28px; font-weight: bold; color: #c2185b;">${claims.length}</div>
                                <div style="color: #c2185b; font-weight: bold; font-size: 12px;">🚨 Sinistres</div>
                            </div>
                            <div style="background: linear-gradient(135deg, #f3e5f5 0%, #e1bee7 100%); padding: 15px; border-radius: 10px; text-align: center; box-shadow: 0 2px 10px rgba(123, 31, 162, 0.2);">
                                <div style="font-size: 28px; font-weight: bold; color: #7b1fa2;">${logs.length}</div>
                                <div style="color: #7b1fa2; font-weight: bold; font-size: 12px;">📝 Autres Actions</div>
                            </div>
                        </div>

                        <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:0 0 25px;">
                            <div style="padding:14px;background:#e3f2fd;border-radius:8px;text-align:center;"><strong style="font-size:22px;color:#1565c0;">${uniqueClients}</strong><br><span style="font-size:11px;color:#455a64;">Dossiers/clients traités</span></div>
                            <div style="padding:14px;background:#e8f5e9;border-radius:8px;text-align:center;"><strong style="font-size:22px;color:#2e7d32;">${estimates.length + policies.length}</strong><br><span style="font-size:11px;color:#455a64;">Actions de production</span></div>
                            <div style="padding:14px;background:#fff3e0;border-radius:8px;text-align:center;"><strong style="font-size:22px;color:#ef6c00;">${tasksCompleted.length}</strong><br><span style="font-size:11px;color:#455a64;">Suivis finalisés</span></div>
                            <div style="padding:14px;background:${remainingWork ? '#ffebee' : '#e8f5e9'};border-radius:8px;text-align:center;"><strong style="font-size:22px;color:${remainingWork ? '#c62828' : '#2e7d32'};">${remainingWork}</strong><br><span style="font-size:11px;color:#455a64;">Éléments restant à traiter</span></div>
                        </div>

                        <!-- Section 1: Emails Envoyés -->
                        <div style="margin-bottom: 30px; border: 1px solid #e3f2fd; border-radius: 10px; overflow: hidden;">
                            <h2 style="background: #1976d2; color: white; margin: 0; padding: 15px 20px; font-size: 16px;">
                                📤 Emails Envoyés (${emailsSent.length})
                            </h2>
                            <div style="padding: 15px;">
                                ${emailsSent.length > 0 ? `
                                    <table style="width: 100%; border-collapse: collapse; font-size: 12px;">
                                        <thead>
                                            <tr style="background: #e3f2fd;">
                                                <th style="padding: 10px; text-align: left; border: 1px solid #bbdefb; width: 100px;">Date</th>
                                                <th style="padding: 10px; text-align: left; border: 1px solid #bbdefb; width: 180px;">Destinataire</th>
                                                <th style="padding: 10px; text-align: left; border: 1px solid #bbdefb; width: 200px;">Objet</th>
                                                <th style="padding: 10px; text-align: left; border: 1px solid #bbdefb;">Contenu</th>
                                                <th style="padding: 10px; text-align: center; border: 1px solid #bbdefb; width: 40px;">PJ</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            ${emailsSent.map((e, idx) => `
                                                <tr>
                                                    <td style="padding: 8px; border: 1px solid #e3f2fd; font-size: 11px; vertical-align: top;">${e.date}</td>
                                                    <td style="padding: 8px; border: 1px solid #e3f2fd; vertical-align: top; word-break: break-word;">
                                                        <strong>${Utils.escapeHtml(e.toEmail)}</strong>
                                                    </td>
                                                    <td style="padding: 8px; border: 1px solid #e3f2fd; font-weight: bold; vertical-align: top; word-break: break-word; max-width: 200px;">
                                                        ${Utils.escapeHtml(e.subject)}
                                                    </td>
                                                    <td style="padding: 8px; border: 1px solid #e3f2fd; vertical-align: top;">
                                                        ${e.body ? `
                                                            <div id="email-preview-${uid}-${idx}" style="color: #666; font-size: 11px;">
                                                                ${Utils.escapeHtml(Utils.truncate(e.body, 150))}
                                                            </div>
                                                            ${e.body.length > 150 ? `
                                                                <div id="email-full-${uid}-${idx}" style="display: none; color: #333; font-size: 11px; white-space: pre-wrap;">
                                                                    ${Utils.escapeHtml(e.body)}
                                                                </div>
                                                                <button onclick="
                                                                    var preview = document.getElementById('email-preview-${uid}-${idx}');
                                                                    var full = document.getElementById('email-full-${uid}-${idx}');
                                                                    if (full.style.display === 'none') {
                                                                        preview.style.display = 'none';
                                                                        full.style.display = 'block';
                                                                        this.textContent = '▲ Réduire';
                                                                    } else {
                                                                        preview.style.display = 'block';
                                                                        full.style.display = 'none';
                                                                        this.textContent = '▼ Voir tout';
                                                                    }
                                                                " style="
                                                                    background: #e3f2fd;
                                                                    border: 1px solid #1976d2;
                                                                    color: #1976d2;
                                                                    padding: 3px 8px;
                                                                    border-radius: 3px;
                                                                    cursor: pointer;
                                                                    font-size: 10px;
                                                                    margin-top: 5px;
                                                                ">▼ Voir tout</button>
                                                            ` : ''}
                                                        ` : '<span style="color: #999;">-</span>'}
                                                    </td>
                                                    <td style="padding: 8px; border: 1px solid #e3f2fd; text-align: center; vertical-align: top;">${e.hasAttachment ? '📎' : '-'}</td>
                                                </tr>
                                            `).join('')}
                                        </tbody>
                                    </table>
                                ` : '<p style="color: #666; font-style: italic; text-align: center; padding: 20px;">Aucun email envoyé aujourd\'hui</p>'}
                            </div>
                        </div>

                        <!-- Section 2: Emails Affectés -->
                        <div style="margin-bottom: 30px; border: 1px solid #e8f5e9; border-radius: 10px; overflow: hidden;">
                            <h2 style="background: #388e3c; color: white; margin: 0; padding: 15px 20px; font-size: 16px;">
                                📥 Emails Affectés (${emailsAffected.length})
                            </h2>
                            <div style="padding: 15px;">
                                ${emailsAffected.length > 0 ? `
                                    <table style="width: 100%; border-collapse: collapse; font-size: 12px;">
                                        <thead>
                                            <tr style="background: #e8f5e9;">
                                                <th style="padding: 10px; text-align: left; border: 1px solid #c8e6c9; width: 100px;">Date</th>
                                                <th style="padding: 10px; text-align: left; border: 1px solid #c8e6c9;">Expéditeur</th>
                                                <th style="padding: 10px; text-align: left; border: 1px solid #c8e6c9;">Objet</th>
                                                <th style="padding: 10px; text-align: left; border: 1px solid #c8e6c9; width: 180px;">Affecté à</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            ${emailsAffected.map(e => `
                                                <tr>
                                                    <td style="padding: 8px; border: 1px solid #e8f5e9; font-size: 11px;">${e.date}</td>
                                                    <td style="padding: 8px; border: 1px solid #e8f5e9;">
                                                        ${Utils.escapeHtml(e.from)}<br>
                                                        <small style="color:#888;">${Utils.escapeHtml(e.fromEmail || '')}</small>
                                                    </td>
                                                    <td style="padding: 8px; border: 1px solid #e8f5e9; font-weight: bold; word-break: break-word;">${Utils.escapeHtml(e.subject)}</td>
                                                    <td style="padding: 8px; border: 1px solid #e8f5e9; color: #388e3c; font-weight: bold;">${Utils.escapeHtml(e.affectedTo)}</td>
                                                </tr>
                                            `).join('')}
                                        </tbody>
                                    </table>
                                ` : '<p style="color: #666; font-style: italic; text-align: center; padding: 20px;">Aucun email affecté aujourd\'hui</p>'}
                            </div>
                        </div>

                        <!-- Section 3: Appels Téléphoniques Aircall -->
                        <div style="margin-bottom: 30px; border: 1px solid #fff8e1; border-radius: 10px; overflow: hidden;">
                            <h2 style="background: #ff8f00; color: white; margin: 0; padding: 15px 20px; font-size: 16px;">
                                📞 Appels Téléphoniques (${(aircallCalls || []).length}) - ${aircallInbound} entrants / ${aircallOutbound} sortants
                            </h2>
                            <div style="padding: 15px;">
                                ${(aircallCalls || []).length > 0 ? `
                                    ${(aircallCalls || []).map((c, idx) => `
                                        <div style="background: #fffbf5; border: 1px solid #ffe0b2; border-radius: 8px; padding: 15px; margin-bottom: 12px;">
                                            <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 10px;">
                                                <div style="display: flex; align-items: center; gap: 12px;">
                                                    <span style="font-size: 24px;">${c.type === 'sortant' ? '📤' : (c.type === 'entrant' ? '📥' : '📞')}</span>
                                                    <div>
                                                        <strong style="color: #ff8f00; font-size: 14px;">${Utils.escapeHtml(c.contact || 'Inconnu')}</strong>
                                                        <br><span style="color: #666; font-size: 12px;">${c.type === 'sortant' ? 'Appel sortant' : (c.type === 'entrant' ? 'Appel entrant' : 'Appel')}</span>
                                                    </div>
                                                </div>
                                                <div style="text-align: right;">
                                                    <span style="background: #ff8f00; color: white; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: bold;">
                                                        ${Utils.escapeHtml(c.duration || '0s')}
                                                    </span>
                                                    <br><span style="color: #888; font-size: 11px; margin-top: 4px; display: inline-block;">
                                                        🕐 ${Utils.escapeHtml(c.time || '')}
                                                    </span>
                                                </div>
                                            </div>
                                            ${c.mood ? `
                                                <div style="margin-bottom: 10px;">
                                                    <span style="
                                                        display: inline-flex;
                                                        align-items: center;
                                                        gap: 6px;
                                                        padding: 4px 10px;
                                                        border-radius: 15px;
                                                        font-size: 12px;
                                                        background: ${c.mood === 'Positif' ? '#e8f5e9' : (c.mood === 'Négatif' ? '#ffebee' : '#f5f5f5')};
                                                        color: ${c.mood === 'Positif' ? '#2e7d32' : (c.mood === 'Négatif' ? '#c62828' : '#666')};
                                                        border: 1px solid ${c.mood === 'Positif' ? '#c8e6c9' : (c.mood === 'Négatif' ? '#ffcdd2' : '#e0e0e0')};
                                                    ">
                                                        ${c.mood === 'Positif' ? '😊' : (c.mood === 'Négatif' ? '😠' : '😐')} ${c.mood}
                                                    </span>
                                                </div>
                                            ` : ''}
                                            ${c.summary ? `
                                                <div style="background: white; border: 1px solid #ffe0b2; border-radius: 6px; padding: 12px; margin-top: 8px;">
                                                    <div style="color: #888; font-size: 10px; margin-bottom: 6px; display: flex; align-items: center; gap: 5px;">
                                                        ✨ Résumé IA
                                                    </div>
                                                    <div id="call-preview-${uid}-${idx}" style="color: #333; font-size: 12px; line-height: 1.5;">
                                                        ${Utils.escapeHtml(Utils.truncate(c.summary, 200))}
                                                    </div>
                                                    ${c.summary.length > 200 ? `
                                                        <div id="call-full-${uid}-${idx}" style="display: none; color: #333; font-size: 12px; line-height: 1.5;">
                                                            ${Utils.escapeHtml(c.summary)}
                                                        </div>
                                                        <button onclick="
                                                            var preview = document.getElementById('call-preview-${uid}-${idx}');
                                                            var full = document.getElementById('call-full-${uid}-${idx}');
                                                            if (full.style.display === 'none') {
                                                                preview.style.display = 'none';
                                                                full.style.display = 'block';
                                                                this.textContent = '▲ Réduire';
                                                            } else {
                                                                preview.style.display = 'block';
                                                                full.style.display = 'none';
                                                                this.textContent = '▼ Voir tout';
                                                            }
                                                        " style="
                                                            background: #fff8e1;
                                                            border: 1px solid #ff8f00;
                                                            color: #ff8f00;
                                                            padding: 3px 8px;
                                                            border-radius: 3px;
                                                            cursor: pointer;
                                                            font-size: 10px;
                                                            margin-top: 8px;
                                                        ">▼ Voir tout</button>
                                                    ` : ''}
                                                </div>
                                            ` : ''}
                                            ${c.topics?.length ? `
                                                <div style="margin-top:8px;font-size:12px;color:#6d4c00;">
                                                    <strong>🏷️ Sujets clés :</strong> ${c.topics.map(topic => Utils.escapeHtml(topic)).join(' · ')}
                                                </div>
                                            ` : ''}
                                            ${c.actionItems?.length ? `
                                                <div style="margin-top:8px;padding:10px;background:#fff3e0;border-radius:6px;font-size:12px;">
                                                    <strong>✅ Actions à entreprendre</strong>
                                                    <ul style="margin:6px 0 0 18px;">${c.actionItems.map(action => `<li>${Utils.escapeHtml(action)}</li>`).join('')}</ul>
                                                </div>
                                            ` : ''}
                                            ${c.transcript ? `
                                                <details style="margin-top:8px;background:white;border:1px solid #ffe0b2;border-radius:6px;padding:10px;">
                                                    <summary style="cursor:pointer;color:#e65100;font-size:12px;font-weight:bold;">🗣️ Voir la transcription complète</summary>
                                                    <div style="margin-top:8px;white-space:pre-wrap;font-size:11px;line-height:1.5;max-height:320px;overflow:auto;">${Utils.escapeHtml(c.transcript)}</div>
                                                </details>
                                            ` : ''}
                                        </div>
                                    `).join('')}
                                ` : '<p style="color: #666; font-style: italic; text-align: center; padding: 20px;">Aucun appel collecté - <a href="https://dashboard.aircall.io/conversations" target="_blank" style="color: #ff8f00;">Ouvrir Aircall</a></p>'}
                            </div>
                        </div>

                        <!-- Section 4: Tâches Terminées -->
                        <div style="margin-bottom: 30px; border: 1px solid #fff3e0; border-radius: 10px; overflow: hidden;">
                            <h2 style="background: #f57c00; color: white; margin: 0; padding: 15px 20px; font-size: 16px;">
                                ✅ Tâches Terminées (${tasksCompleted.length})
                            </h2>
                            <div style="padding: 15px;">
                                ${tasksCompleted.length > 0 ? `
                                    ${tasksCompleted.map((t, idx) => `
                                        <div style="background: #fff8f0; border: 1px solid #ffe0b2; border-radius: 8px; padding: 15px; margin-bottom: 10px;">
                                            <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 10px;">
                                                <div>
                                                    <strong style="color: #f57c00; font-size: 14px;">${t.isPriority ? '🔴 ' : ''}${Utils.escapeHtml(t.title)}</strong>
                                                    <br><span style="color: #666; font-size: 12px;">Client: ${Utils.escapeHtml(t.client)}</span>
                                                </div>
                                                <span style="background: #f57c00; color: white; padding: 3px 10px; border-radius: 12px; font-size: 11px;">
                                                    ${t.closedTime ? `⏰ ${t.closedTime}` : `Terminée le ${t.completedDate}`}
                                                </span>
                                            </div>
                                            ${t.content ? `
                                                <div style="background: white; border: 1px solid #ffe0b2; border-radius: 5px; padding: 10px; margin-top: 10px;">
                                                    <div style="color: #888; font-size: 10px; margin-bottom: 5px;">📝 Contenu de la tâche:</div>
                                                    <div id="task-preview-${uid}-${idx}" style="color: #333; font-size: 12px;">
                                                        ${Utils.escapeHtml(Utils.truncate(t.content, 200))}
                                                    </div>
                                                    ${t.content.length > 200 ? `
                                                        <div id="task-full-${uid}-${idx}" style="display: none; color: #333; font-size: 12px; white-space: pre-wrap;">
                                                            ${Utils.escapeHtml(t.content)}
                                                        </div>
                                                        <button onclick="
                                                            var preview = document.getElementById('task-preview-${uid}-${idx}');
                                                            var full = document.getElementById('task-full-${uid}-${idx}');
                                                            if (full.style.display === 'none') {
                                                                preview.style.display = 'none';
                                                                full.style.display = 'block';
                                                                this.textContent = '▲ Réduire';
                                                            } else {
                                                                preview.style.display = 'block';
                                                                full.style.display = 'none';
                                                                this.textContent = '▼ Voir tout';
                                                            }
                                                        " style="
                                                            background: #fff3e0;
                                                            border: 1px solid #f57c00;
                                                            color: #f57c00;
                                                            padding: 3px 8px;
                                                            border-radius: 3px;
                                                            cursor: pointer;
                                                            font-size: 10px;
                                                            margin-top: 5px;
                                                        ">▼ Voir tout</button>
                                                    ` : ''}
                                                </div>
                                            ` : ''}
                                            <div style="color: #999; font-size: 11px; margin-top: 8px;">
                                                Créée par ${Utils.escapeHtml(t.createdBy)} le ${t.createdDate}
                                            </div>
                                        </div>
                                    `).join('')}
                                ` : '<p style="color: #666; font-style: italic; text-align: center; padding: 20px;">Aucune tâche terminée aujourd\'hui</p>'}
                            </div>
                        </div>

                        <!-- Section 4: Tâches en Retard -->
                        <div style="margin-bottom: 30px; border: 1px solid #ffebee; border-radius: 10px; overflow: hidden;">
                            <h2 style="background: #d32f2f; color: white; margin: 0; padding: 15px 20px; font-size: 16px;">
                                ⚠️ Tâches en Retard (${tasksOverdue.length})
                            </h2>
                            <div style="padding: 15px;">
                                ${tasksOverdue.length > 0 ? `
                                    ${tasksOverdue.map((t, idx) => `
                                        <div style="background: ${t.daysOverdue > 30 ? '#ffcdd2' : t.daysOverdue > 7 ? '#ffe0b2' : '#fff8e1'}; border: 1px solid ${t.daysOverdue > 30 ? '#ef9a9a' : t.daysOverdue > 7 ? '#ffcc80' : '#fff59d'}; border-radius: 8px; padding: 15px; margin-bottom: 10px;">
                                            <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 10px;">
                                                <div>
                                                    <strong style="color: ${t.daysOverdue > 30 ? '#b71c1c' : t.daysOverdue > 7 ? '#e65100' : '#f57c00'}; font-size: 14px;">
                                                        ${t.isPriority ? '🔴 ' : ''}${Utils.escapeHtml(t.title)}
                                                    </strong>
                                                    <br><span style="color: #666; font-size: 12px;">Client: ${Utils.escapeHtml(t.client)}</span>
                                                </div>
                                                <div style="text-align: right;">
                                                    <span style="background: ${t.daysOverdue > 30 ? '#b71c1c' : t.daysOverdue > 7 ? '#e65100' : '#f57c00'}; color: white; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: bold;">
                                                        ${t.daysOverdue} jour${t.daysOverdue > 1 ? 's' : ''} de retard
                                                    </span>
                                                    <br><span style="color: #666; font-size: 11px; margin-top: 3px; display: inline-block;">Échéance: ${t.dueDate}</span>
                                                </div>
                                            </div>
                                            ${t.content ? `
                                                <div style="background: rgba(255,255,255,0.7); border-radius: 5px; padding: 10px; margin-top: 10px;">
                                                    <div style="color: #888; font-size: 10px; margin-bottom: 5px;">📝 Contenu de la tâche:</div>
                                                    <div id="overdue-preview-${uid}-${idx}" style="color: #333; font-size: 12px;">
                                                        ${Utils.escapeHtml(Utils.truncate(t.content, 200))}
                                                    </div>
                                                    ${t.content.length > 200 ? `
                                                        <div id="overdue-full-${uid}-${idx}" style="display: none; color: #333; font-size: 12px; white-space: pre-wrap;">
                                                            ${Utils.escapeHtml(t.content)}
                                                        </div>
                                                        <button onclick="
                                                            var preview = document.getElementById('overdue-preview-${uid}-${idx}');
                                                            var full = document.getElementById('overdue-full-${uid}-${idx}');
                                                            if (full.style.display === 'none') {
                                                                preview.style.display = 'none';
                                                                full.style.display = 'block';
                                                                this.textContent = '▲ Réduire';
                                                            } else {
                                                                preview.style.display = 'block';
                                                                full.style.display = 'none';
                                                                this.textContent = '▼ Voir tout';
                                                            }
                                                        " style="
                                                            background: rgba(255,255,255,0.8);
                                                            border: 1px solid #d32f2f;
                                                            color: #d32f2f;
                                                            padding: 3px 8px;
                                                            border-radius: 3px;
                                                            cursor: pointer;
                                                            font-size: 10px;
                                                            margin-top: 5px;
                                                        ">▼ Voir tout</button>
                                                    ` : ''}
                                                </div>
                                            ` : ''}
                                        </div>
                                    `).join('')}
                                ` : '<p style="color: #388e3c; text-align: center; padding: 20px; font-weight: bold;">🎉 Aucune tâche en retard ! Bravo !</p>'}
                            </div>
                        </div>

                        <!-- Section 5: Devis -->
                        <div style="margin-bottom: 30px; border: 1px solid #e0f7fa; border-radius: 10px; overflow: hidden;">
                            <h2 style="background: #0097a7; color: white; margin: 0; padding: 15px 20px; font-size: 16px;">
                                📋 Devis (${estimates.length})
                            </h2>
                            <div style="padding: 15px;">
                                ${estimates.length > 0 ? `
                                    ${estimates.map(log => `
                                        <div style="background: #e0f7fa; border: 1px solid #b2ebf2; border-radius: 8px; padding: 15px; margin-bottom: 10px;">
                                            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                                                <strong style="color: #0097a7; font-size: 14px;">
                                                    ${log.action}
                                                </strong>
                                                <span style="color: #00acc1; font-size: 12px;">${log.date}</span>
                                            </div>
                                            <div style="margin-bottom: 10px; color: #333; font-size: 13px;">
                                                <strong>Devis:</strong> ${Utils.escapeHtml(log.entityName)}
                                                <span style="color: #999; font-size: 11px;">(N° ${log.entityId})</span>
                                            </div>
                                            ${log.changes.length > 0 ? `
                                                <table style="width: 100%; border-collapse: collapse; font-size: 12px; background: white; border-radius: 5px; overflow: hidden;">
                                                    <thead>
                                                        <tr>
                                                            <th style="padding: 8px; text-align: left; background: #b2ebf2; border: 1px solid #80deea;">Champ</th>
                                                            <th style="padding: 8px; text-align: left; background: #ffcdd2; border: 1px solid #ef9a9a;">Avant</th>
                                                            <th style="padding: 8px; text-align: left; background: #c8e6c9; border: 1px solid #a5d6a7;">Après</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        ${log.changes.map(c => `
                                                            <tr>
                                                                <td style="padding: 6px; border: 1px solid #b2ebf2; font-weight: bold;">${Utils.escapeHtml(c.field)}</td>
                                                                <td style="padding: 6px; border: 1px solid #ffcdd2; color: #c62828;">${Utils.escapeHtml(c.oldValue)}</td>
                                                                <td style="padding: 6px; border: 1px solid #c8e6c9; color: #2e7d32;">${Utils.escapeHtml(c.newValue)}</td>
                                                            </tr>
                                                        `).join('')}
                                                    </tbody>
                                                </table>
                                            ` : '<p style="color: #666; font-size: 12px; margin: 0;">Création du devis</p>'}
                                        </div>
                                    `).join('')}
                                ` : '<p style="color: #666; font-style: italic; text-align: center; padding: 20px;">Aucun devis créé ou modifié aujourd\'hui</p>'}
                            </div>
                        </div>

                        <!-- Section 6: Contrats -->
                        <div style="margin-bottom: 30px; border: 1px solid #e8eaf6; border-radius: 10px; overflow: hidden;">
                            <h2 style="background: #3f51b5; color: white; margin: 0; padding: 15px 20px; font-size: 16px;">
                                📄 Contrats (${policies.length})
                            </h2>
                            <div style="padding: 15px;">
                                ${policies.length > 0 ? `
                                    ${policies.map(log => `
                                        <div style="background: #e8eaf6; border: 1px solid #c5cae9; border-radius: 8px; padding: 15px; margin-bottom: 10px;">
                                            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                                                <strong style="color: #3f51b5; font-size: 14px;">
                                                    ${log.action}
                                                </strong>
                                                <span style="color: #5c6bc0; font-size: 12px;">${log.date}</span>
                                            </div>
                                            <div style="margin-bottom: 10px; color: #333; font-size: 13px;">
                                                <strong>Contrat:</strong> ${Utils.escapeHtml(log.entityName)}
                                                <span style="color: #999; font-size: 11px;">(N° ${log.entityId})</span>
                                            </div>
                                            ${log.changes.length > 0 ? `
                                                <table style="width: 100%; border-collapse: collapse; font-size: 12px; background: white; border-radius: 5px; overflow: hidden;">
                                                    <thead>
                                                        <tr>
                                                            <th style="padding: 8px; text-align: left; background: #c5cae9; border: 1px solid #9fa8da;">Champ</th>
                                                            <th style="padding: 8px; text-align: left; background: #ffcdd2; border: 1px solid #ef9a9a;">Avant</th>
                                                            <th style="padding: 8px; text-align: left; background: #c8e6c9; border: 1px solid #a5d6a7;">Après</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        ${log.changes.map(c => `
                                                            <tr>
                                                                <td style="padding: 6px; border: 1px solid #c5cae9; font-weight: bold;">${Utils.escapeHtml(c.field)}</td>
                                                                <td style="padding: 6px; border: 1px solid #ffcdd2; color: #c62828;">${Utils.escapeHtml(c.oldValue)}</td>
                                                                <td style="padding: 6px; border: 1px solid #c8e6c9; color: #2e7d32;">${Utils.escapeHtml(c.newValue)}</td>
                                                            </tr>
                                                        `).join('')}
                                                    </tbody>
                                                </table>
                                            ` : '<p style="color: #666; font-size: 12px; margin: 0;">Création du contrat</p>'}
                                        </div>
                                    `).join('')}
                                ` : '<p style="color: #666; font-style: italic; text-align: center; padding: 20px;">Aucun contrat créé ou modifié aujourd\'hui</p>'}
                            </div>
                        </div>

                        <!-- Section 7: Sinistres -->
                        <div style="margin-bottom: 30px; border: 1px solid #fce4ec; border-radius: 10px; overflow: hidden;">
                            <h2 style="background: #c2185b; color: white; margin: 0; padding: 15px 20px; font-size: 16px;">
                                🚨 Sinistres (${claims.length})
                            </h2>
                            <div style="padding: 15px;">
                                ${claims.length > 0 ? `
                                    ${claims.map(log => `
                                        <div style="background: #fce4ec; border: 1px solid #f8bbd9; border-radius: 8px; padding: 15px; margin-bottom: 10px;">
                                            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                                                <strong style="color: #c2185b; font-size: 14px;">
                                                    ${log.action}
                                                </strong>
                                                <span style="color: #d81b60; font-size: 12px;">${log.date}</span>
                                            </div>
                                            <div style="margin-bottom: 10px; color: #333; font-size: 13px;">
                                                <strong>Sinistre:</strong> ${Utils.escapeHtml(log.entityName)}
                                                <span style="color: #999; font-size: 11px;">(N° ${log.entityId})</span>
                                            </div>
                                            ${log.changes.length > 0 ? `
                                                <table style="width: 100%; border-collapse: collapse; font-size: 12px; background: white; border-radius: 5px; overflow: hidden;">
                                                    <thead>
                                                        <tr>
                                                            <th style="padding: 8px; text-align: left; background: #f8bbd9; border: 1px solid #f48fb1;">Champ</th>
                                                            <th style="padding: 8px; text-align: left; background: #ffcdd2; border: 1px solid #ef9a9a;">Avant</th>
                                                            <th style="padding: 8px; text-align: left; background: #c8e6c9; border: 1px solid #a5d6a7;">Après</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        ${log.changes.map(c => `
                                                            <tr>
                                                                <td style="padding: 6px; border: 1px solid #f8bbd9; font-weight: bold;">${Utils.escapeHtml(c.field)}</td>
                                                                <td style="padding: 6px; border: 1px solid #ffcdd2; color: #c62828;">${Utils.escapeHtml(c.oldValue)}</td>
                                                                <td style="padding: 6px; border: 1px solid #c8e6c9; color: #2e7d32;">${Utils.escapeHtml(c.newValue)}</td>
                                                            </tr>
                                                        `).join('')}
                                                    </tbody>
                                                </table>
                                            ` : '<p style="color: #666; font-size: 12px; margin: 0;">Création du sinistre</p>'}
                                        </div>
                                    `).join('')}
                                ` : '<p style="color: #666; font-style: italic; text-align: center; padding: 20px;">Aucun sinistre créé ou modifié aujourd\'hui</p>'}
                            </div>
                        </div>

                        <!-- Section 8: Journalisation (Vulgarisée) - RÉDUIT PAR DÉFAUT -->
                        <div style="margin-bottom: 30px; border: 1px solid #f3e5f5; border-radius: 10px; overflow: hidden;">
                            <details>
                                <summary style="background: #7b1fa2; color: white; margin: 0; padding: 15px 20px; font-size: 16px; cursor: pointer; list-style: none; display: flex; justify-content: space-between; align-items: center; user-select: none;">
                                    <span>📝 Actions sur Fiches Clients (${logs.length})</span>
                                    <span style="font-size: 11px; background: rgba(255,255,255,0.2); padding: 4px 12px; border-radius: 15px;">▶ Cliquer pour voir</span>
                                </summary>
                                <div style="padding: 15px; max-height: 600px; overflow-y: auto;">
                                ${logs.length > 0 ? `
                                    ${logs.map(log => {
                                        // Vulgariser l'entrée
                                        const vulgarized = LogVulgarizer.vulgarize(log);

                                        // Générer le lien vers l'entité
                                        let entityLink = '#';
                                        let entityIcon = '📄';
                                        if (log.tableRaw) {
                                            const tableType = log.tableRaw.toLowerCase();
                                            if (tableType.includes('client')) {
                                                entityLink = `https://courtage.modulr.fr/fr/scripts/clients/clients_card.php?id=${log.entityId}`;
                                                entityIcon = '👤';
                                            } else if (tableType.includes('task') || tableType.includes('tâche')) {
                                                entityLink = `https://courtage.modulr.fr/fr/scripts/Tasks/TasksCard.php?id=${log.entityId}`;
                                                entityIcon = '✅';
                                            } else if (tableType.includes('email')) {
                                                entityLink = `https://courtage.modulr.fr/fr/scripts/sent_emails/sent_emails_view.php?id=${log.entityId}`;
                                                entityIcon = '📧';
                                            } else if (tableType.includes('estimate') || tableType.includes('devis')) {
                                                entityLink = `https://courtage.modulr.fr/fr/scripts/estimates/estimates_card.php?id=${log.entityId}`;
                                                entityIcon = '📋';
                                            } else if (tableType.includes('polic') || tableType.includes('contrat')) {
                                                entityLink = `https://courtage.modulr.fr/fr/scripts/policies/policies_card.php?id=${log.entityId}`;
                                                entityIcon = '📄';
                                            } else if (tableType.includes('claim') || tableType.includes('sinistre')) {
                                                entityLink = `https://courtage.modulr.fr/fr/scripts/claims/claims_card.php?id=${log.entityId}`;
                                                entityIcon = '🚨';
                                            }
                                        }

                                        return `
                                        <div style="background: #faf5fc; border: 1px solid #e1bee7; border-radius: 8px; padding: 15px; margin-bottom: 15px;">
                                            <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 8px;">
                                                <div>
                                                    <strong style="color: #7b1fa2; font-size: 14px;">
                                                        ${vulgarized.title}
                                                    </strong>
                                                    ${vulgarized.summary ? `<br><span style="color: #666; font-size: 12px;">${vulgarized.summary}</span>` : ''}
                                                </div>
                                                <span style="color: #9575cd; font-size: 11px; white-space: nowrap;">${log.date}</span>
                                            </div>
                                            <div style="margin-bottom: 10px;">
                                                <a href="${entityLink}" target="_blank" style="color: #7b1fa2; text-decoration: none; font-size: 13px;">
                                                    ${entityIcon} ${Utils.escapeHtml(log.entityName || 'Voir la fiche')}
                                                </a>
                                            </div>
                                            ${vulgarized.details && vulgarized.details.length > 0 ? `
                                                <details style="margin-top: 10px;">
                                                    <summary style="cursor: pointer; color: #7b1fa2; font-size: 12px; padding: 5px 0;">
                                                        📋 Voir les ${vulgarized.details.length} modification(s)
                                                    </summary>
                                                    <table style="width: 100%; border-collapse: collapse; font-size: 11px; background: white; border-radius: 5px; overflow: hidden; margin-top: 8px;">
                                                        <thead>
                                                            <tr>
                                                                <th style="padding: 6px; text-align: left; background: #e1bee7; border: 1px solid #ce93d8; width: 30%;">Champ</th>
                                                                <th style="padding: 6px; text-align: left; background: #ffcdd2; border: 1px solid #ef9a9a; width: 35%;">Avant</th>
                                                                <th style="padding: 6px; text-align: left; background: #c8e6c9; border: 1px solid #a5d6a7; width: 35%;">Après</th>
                                                            </tr>
                                                        </thead>
                                                        <tbody>
                                                            ${vulgarized.details.map(c => `
                                                                <tr>
                                                                    <td style="padding: 5px; border: 1px solid #e1bee7; font-weight: bold;">${Utils.escapeHtml(c.field)}</td>
                                                                    <td style="padding: 5px; border: 1px solid #ffcdd2; color: #c62828;">${Utils.escapeHtml(String(c.oldValue || '-'))}</td>
                                                                    <td style="padding: 5px; border: 1px solid #c8e6c9; color: #2e7d32;">${Utils.escapeHtml(String(c.newValue || '-'))}</td>
                                                                </tr>
                                                            `).join('')}
                                                        </tbody>
                                                    </table>
                                                </details>
                                            ` : ''}
                                        </div>
                                    `;
                                    }).join('')}
                                ` : '<p style="color: #666; font-style: italic; text-align: center; padding: 20px;">Aucune action sur les fiches aujourd\'hui</p>'}
                                </div>
                            </details>
                        </div>

                        ${notes ? `
                        <div style="margin:25px 0;padding:18px;background:#fffde7;border-left:5px solid #f9a825;border-radius:8px;">
                            <strong style="color:#6d4c00;">🗒️ Notes et précisions</strong>
                            <div style="margin-top:10px;white-space:pre-wrap;color:#4e342e;line-height:1.55;">${Utils.escapeHtml(notes)}</div>
                        </div>` : ''}

                        <!-- Footer -->
                        <div style="text-align: center; color: #999; padding-top: 20px; border-top: 1px solid #eee; font-size: 12px;">
                            <p>📊 Rapport généré automatiquement par LTOA Modulr Script v4</p>
                            <p>${new Date().toLocaleString('fr-FR')}</p>
                        </div>
                    </div>
                </div>
            `;
        },

        show() {
            const existing = document.getElementById('ltoa-report-modal');
            if (existing) existing.remove();

            document.body.insertAdjacentHTML('beforeend', this.generateHTML());

            document.getElementById('ltoa-close-report').addEventListener('click', () => {
                document.getElementById('ltoa-report-modal').remove();
            });

            document.getElementById('ltoa-view-by-client').addEventListener('click', () => this.showByClientView());
            document.getElementById('ltoa-export-html').addEventListener('click', () => this.exportHTML());
            document.getElementById('ltoa-view-chrono').addEventListener('click', () => this.showChronoView());

            document.addEventListener('keydown', (e) => {
                if (e.key === 'Escape') {
                    const modal = document.getElementById('ltoa-report-modal');
                    if (modal) modal.remove();
                }
            });
        },

        // ============================================
        // VUE PAR CLIENT
        // ============================================
        showByClientView() {
            try {
                const { emailsSent, emailsAffected, aircallCalls, tasksCompleted, tasksOverdue, logs, estimates, policies, claims, user, date, clientIndex } = this.data;

                // Vérifier si c'est un rapport pour un jour passé
                const realToday = Utils.getRealTodayDate();
                const isPastDate = date !== realToday;

                // Regrouper toutes les données par client (utiliser l'ID client comme clé si disponible)
                const clientsMap = new Map();

                // Helper pour ajouter une entrée à un client
                const addToClient = (clientName, clientId, clientEmail, type, item) => {
                    // Utiliser l'ID client comme clé principale si disponible
                    let key = clientId ? `id_${clientId}` : (clientName || 'Sans client associé');

                    if (!clientName || clientName === 'N/A' || clientName === 'Non associé') {
                        if (!clientId) {
                            key = 'Sans client associé';
                            clientName = 'Sans client associé';
                        }
                    }

                    if (!clientsMap.has(key)) {
                        clientsMap.set(key, {
                            name: clientName || 'Client inconnu',
                            id: clientId,
                            email: clientEmail,
                            emailsSent: [],
                            emailsAffected: [],
                            aircallCalls: [],
                            tasksCompleted: [],
                            tasksOverdue: [],
                            estimates: [],
                            policies: [],
                            claims: [],
                            logs: []
                        });
                    }

                    // Mettre à jour les infos si on a de meilleures données
                    const client = clientsMap.get(key);
                    if (clientId && !client.id) client.id = clientId;
                    if (clientEmail && !client.email) client.email = clientEmail;
                    if (clientName && clientName !== 'Sans client associé' && client.name === 'Client inconnu') {
                        client.name = clientName;
                    }

                    client[type].push(item);
                };

                // Emails envoyés - utiliser les données enrichies si disponibles
                emailsSent.forEach(e => {
                    const clientName = e.clientName || e.recipientName || e.to || null;
                    const clientId = e.clientId || null;
                    const clientEmail = e.toEmail || null;
                    addToClient(clientName, clientId, clientEmail, 'emailsSent', e);
                });

                // Emails affectés - utiliser les données enrichies si disponibles
                emailsAffected.forEach(e => {
                    const clientName = e.clientName || e.fromName || e.from || null;
                    const clientId = e.clientId || null;
                    const clientEmail = e.fromEmail || null;
                    addToClient(clientName, clientId, clientEmail, 'emailsAffected', e);
                });

                // Appels Aircall - regrouper par nom de contact
                // Essayer de matcher avec les clients existants (nom entre parenthèses, etc.)
                (aircallCalls || []).forEach(call => {
                    const contactName = call.contact || 'Contact inconnu';
                    const contactLower = contactName.toLowerCase().trim();

                    // Chercher un client existant qui correspond
                    let matchedKey = null;
                    clientsMap.forEach((client, key) => {
                        const clientNameLower = client.name.toLowerCase();
                        // Match direct
                        if (clientNameLower.includes(contactLower) || contactLower.includes(clientNameLower)) {
                            matchedKey = key;
                        }
                        // Match avec nom entre parenthèses (ex: "MAUD'AUTO (maude mancini)")
                        const parenMatch = clientNameLower.match(/\(([^)]+)\)/);
                        if (parenMatch) {
                            const nameInParen = parenMatch[1].toLowerCase().trim();
                            if (nameInParen.includes(contactLower) || contactLower.includes(nameInParen)) {
                                matchedKey = key;
                            }
                        }
                        // Match par parties du nom
                        if (!matchedKey) {
                            const contactParts = contactLower.split(/[\s,]+/).filter(p => p.length > 2);
                            const clientParts = clientNameLower.split(/[\s,()]+/).filter(p => p.length > 2);
                            let matchCount = 0;
                            for (const cp of contactParts) {
                                if (clientParts.some(clp => clp.includes(cp) || cp.includes(clp))) {
                                    matchCount++;
                                }
                            }
                            // Si au moins 2 parties matchent ou toutes les parties du contact matchent
                            if (matchCount >= 2 || (contactParts.length > 0 && matchCount === contactParts.length)) {
                                matchedKey = key;
                            }
                        }
                    });

                    if (matchedKey) {
                        clientsMap.get(matchedKey).aircallCalls.push(call);
                        Utils.log(`Appel "${contactName}" associé au client "${clientsMap.get(matchedKey).name}"`);
                    } else {
                        addToClient(contactName, null, null, 'aircallCalls', call);
                    }
                });

                // Tâches terminées
                tasksCompleted.forEach(t => {
                    addToClient(t.client, t.clientId, null, 'tasksCompleted', t);
                });

                // Tâches en retard
                tasksOverdue.forEach(t => {
                    addToClient(t.client, t.clientId, null, 'tasksOverdue', t);
                });

                // Devis - utiliser les données enrichies
                estimates.forEach(e => {
                    const clientName = e.clientName || e.entityName;
                    addToClient(clientName, e.clientId, e.clientEmail, 'estimates', e);
                });

                // Contrats - utiliser les données enrichies
                policies.forEach(p => {
                    const clientName = p.clientName || p.entityName;
                    addToClient(clientName, p.clientId, p.clientEmail, 'policies', p);
                });

                // Sinistres - utiliser les données enrichies
                claims.forEach(c => {
                    const clientName = c.clientName || c.entityName;
                    addToClient(clientName, c.clientId, c.clientEmail, 'claims', c);
                });

                // Logs - regrouper par client si l'ID client est présent dans les changements
                logs.forEach(l => {
                    // Exclure les tables qui ne sont pas des clients
                    const tableRawLower = (l.tableRaw || '').toLowerCase();
                    if (tableRawLower.includes('utilisateur') || tableRawLower.includes('user') ||
                        tableRawLower.includes('collaborateur') || tableRawLower.includes('employe')) {
                        return; // Skip - pas un client
                    }

                    // Chercher un ID client dans les changements ou dans la table
                    let clientId = l.clientId;
                    let clientName = l.entityName;

                    // Si c'est une table client, utiliser l'entityId
                    if (l.tableRaw && tableRawLower.includes('client')) {
                        clientId = l.entityId;
                    }

                    // Chercher dans les changements si y'a un champ client_id ou Client (ID)
                    if (l.changes && Array.isArray(l.changes)) {
                        for (const change of l.changes) {
                            const fieldRaw = (change.fieldRaw || '').toLowerCase();
                            const fieldName = (change.field || '').toLowerCase();
                            // Chercher client_id, Client (ID), etc.
                            if (fieldRaw === 'client_id' || fieldRaw.includes('client') && fieldRaw.includes('id') ||
                                fieldName.includes('client') && fieldName.includes('id')) {
                                const val = change.newValueRaw || change.newValue || change.oldValueRaw || change.oldValue;
                                if (val && /^\d+$/.test(String(val).trim())) {
                                    clientId = String(val).trim();
                                    Utils.log(`Log "${l.actionRaw}" sur ${l.tableRaw}: Client ID trouvé = ${clientId}`);
                                    break;
                                }
                            }
                        }
                    }

                    if (clientId) {
                        addToClient(clientName, clientId, l.clientEmail, 'logs', l);
                    }
                });

                // Fusionner les clients qui ont le même ID
                const mergedClients = new Map();
                clientsMap.forEach((client, key) => {
                    if (client.id) {
                        const existingKey = `id_${client.id}`;
                        if (mergedClients.has(existingKey)) {
                            const existing = mergedClients.get(existingKey);
                            // Fusionner les données
                            existing.emailsSent.push(...client.emailsSent);
                            existing.emailsAffected.push(...client.emailsAffected);
                            existing.aircallCalls.push(...client.aircallCalls);
                            existing.tasksCompleted.push(...client.tasksCompleted);
                            existing.tasksOverdue.push(...client.tasksOverdue);
                            existing.estimates.push(...client.estimates);
                            existing.policies.push(...client.policies);
                            existing.claims.push(...client.claims);
                            existing.logs.push(...client.logs);
                            // Garder le meilleur nom/email
                            if (!existing.email && client.email) existing.email = client.email;
                            if (client.name && client.name !== 'Sans client associé') existing.name = client.name;
                        } else {
                            mergedClients.set(existingKey, client);
                        }
                    } else {
                        mergedClients.set(key, client);
                    }
                });

                // Trier les clients par nombre d'actions (plus actifs en premier)
                const sortedClients = Array.from(mergedClients.values()).sort((a, b) => {
                    const countA = a.emailsSent.length + a.emailsAffected.length + a.aircallCalls.length + a.tasksCompleted.length +
                                  a.estimates.length + a.policies.length + a.claims.length + a.logs.length;
                    const countB = b.emailsSent.length + b.emailsAffected.length + b.aircallCalls.length + b.tasksCompleted.length +
                                  b.estimates.length + b.policies.length + b.claims.length + b.logs.length;
                    return countB - countA;
                });

                // Filtrer les clients sans aucune action UTILE (exclure ceux qui n'ont que des tâches en retard)
                const activeClients = sortedClients.filter(c => {
                    const hasUsefulActions = c.emailsSent.length + c.emailsAffected.length + c.aircallCalls.length +
                        c.tasksCompleted.length + c.estimates.length + c.policies.length + c.claims.length + c.logs.length > 0;
                    // Si le client n'a que des tâches en retard et rien d'autre, on l'exclut
                    if (!hasUsefulActions && c.tasksOverdue.length > 0) {
                        return false;
                    }
                    return hasUsefulActions || c.tasksOverdue.length > 0;
                });

            // Générer le HTML de la vue par client
            const clientViewHTML = `
                <div id="ltoa-client-view-modal" style="
                    position: fixed;
                    top: 0;
                    left: 0;
                    width: 100%;
                    height: 100%;
                    background: rgba(0,0,0,0.9);
                    z-index: 1000000;
                    overflow-y: auto;
                    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                ">
                    <div style="
                        max-width: 1100px;
                        margin: 20px auto;
                        background: #f5f5f5;
                        border-radius: 15px;
                        padding: 25px;
                        box-shadow: 0 10px 50px rgba(0,0,0,0.3);
                    ">
                        <!-- Header -->
                        <div style="display: flex; justify-content: space-between; align-items: center; background: white; padding: 20px; border-radius: 12px; margin-bottom: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.05);">
                            <div>
                                <h1 style="color: #1565c0; margin: 0; font-size: 22px; font-weight: 600;">👤 Vue par Client</h1>
                                <p style="color: #666; margin: 5px 0 0 0; font-size: 14px;">
                                    <strong>${user}</strong> • ${date} • ${activeClients.length} clients
                                </p>
                            </div>
                            <button id="ltoa-close-client-view" style="
                                background: linear-gradient(135deg, #666, #444);
                                color: white;
                                border: none;
                                padding: 12px 24px;
                                border-radius: 8px;
                                cursor: pointer;
                                font-size: 13px;
                                font-weight: bold;
                            ">✕ Fermer</button>
                        </div>

                        <!-- Liste des clients -->
                        ${activeClients.length > 0 ? activeClients.map((client, clientIdx) => {
                            const clientLink = client.id ?
                                `https://courtage.modulr.fr/fr/scripts/clients/clients_card.php?id=${client.id}` : '#';
                            const cuid = 'c' + clientIdx + '_' + Date.now();

                            return `
                            <div style="background: white; border-radius: 12px; margin-bottom: 20px; overflow: hidden; box-shadow: 0 2px 10px rgba(0,0,0,0.08);">
                                <!-- En-tête client avec gradient -->
                                <div style="background: linear-gradient(135deg, #1565c0 0%, #0d47a1 100%); color: white; padding: 18px 22px;">
                                    <div style="display: flex; justify-content: space-between; align-items: flex-start;">
                                        <div>
                                            <a href="${clientLink}" target="_blank" style="color: white; text-decoration: none; font-size: 18px; font-weight: 600;">
                                                👤 ${Utils.escapeHtml(client.name)}
                                            </a>
                                            ${client.id ? `<span style="background: rgba(255,255,255,0.2); padding: 3px 10px; border-radius: 12px; font-size: 11px; margin-left: 10px;">N° ${client.id}</span>` : ''}
                                            ${client.email ? `<div style="opacity: 0.8; font-size: 12px; margin-top: 5px;">📧 ${Utils.escapeHtml(client.email)}</div>` : ''}
                                        </div>
                                        <div style="display: flex; gap: 6px; flex-wrap: wrap;">
                                            ${client.emailsSent.length > 0 ? `<span style="background: #2196f3; padding: 4px 12px; border-radius: 15px; font-size: 11px; font-weight: 500;">📤 ${client.emailsSent.length}</span>` : ''}
                                            ${client.emailsAffected.length > 0 ? `<span style="background: #4caf50; padding: 4px 12px; border-radius: 15px; font-size: 11px; font-weight: 500;">📥 ${client.emailsAffected.length}</span>` : ''}
                                            ${client.aircallCalls.length > 0 ? `<span style="background: #ff9800; padding: 4px 12px; border-radius: 15px; font-size: 11px; font-weight: 500;">📞 ${client.aircallCalls.length}</span>` : ''}
                                            ${client.tasksCompleted.length > 0 ? `<span style="background: #ff5722; padding: 4px 12px; border-radius: 15px; font-size: 11px; font-weight: 500;">✅ ${client.tasksCompleted.length}</span>` : ''}
                                        </div>
                                    </div>
                                </div>

                                <!-- Contenu avec cartes colorées -->
                                <div style="padding: 18px; display: grid; gap: 12px;">

                                    ${client.emailsSent.length > 0 ? `
                                    <!-- Emails envoyés -->
                                    <div style="background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%); border-radius: 10px; padding: 15px; border-left: 4px solid #1976d2;">
                                        <div style="font-weight: 600; color: #1565c0; margin-bottom: 10px; font-size: 14px;">📤 Emails envoyés (${client.emailsSent.length})</div>
                                        ${client.emailsSent.map((e, eIdx) => `
                                            <div style="background: white; border-radius: 6px; padding: 10px; margin-bottom: 6px;">
                                                <div style="display: flex; justify-content: space-between;">
                                                    <strong style="color: #333; font-size: 13px;">${Utils.escapeHtml(e.subject || 'Sans objet')}</strong>
                                                    <span style="color: #1976d2; font-size: 11px; font-weight: 500;">${e.time || ''}</span>
                                                </div>
                                                ${e.body ? `
                                                    <div id="email_short_${cuid}_${eIdx}" style="color: #666; font-size: 12px; margin-top: 6px; line-height: 1.4;">${Utils.escapeHtml(Utils.truncate(e.body, 150))}</div>
                                                    ${e.body.length > 150 ? `
                                                        <div id="email_full_${cuid}_${eIdx}" style="display: none; color: #666; font-size: 12px; margin-top: 6px; line-height: 1.4; white-space: pre-wrap;">${Utils.escapeHtml(e.body)}</div>
                                                        <button onclick="var s=document.getElementById('email_short_${cuid}_${eIdx}');var f=document.getElementById('email_full_${cuid}_${eIdx}');if(f.style.display==='none'){f.style.display='block';s.style.display='none';this.textContent='▲ Réduire';}else{f.style.display='none';s.style.display='block';this.textContent='▼ Voir plus';}" style="background: #1976d2; color: white; border: none; padding: 4px 10px; border-radius: 4px; font-size: 11px; cursor: pointer; margin-top: 6px;">▼ Voir plus</button>
                                                    ` : ''}
                                                ` : ''}
                                            </div>
                                        `).join('')}
                                    </div>
                                    ` : ''}

                                    ${client.emailsAffected.length > 0 ? `
                                    <!-- Emails reçus -->
                                    <div style="background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%); border-radius: 10px; padding: 15px; border-left: 4px solid #388e3c;">
                                        <div style="font-weight: 600; color: #2e7d32; margin-bottom: 10px; font-size: 14px;">📥 Emails reçus/affectés (${client.emailsAffected.length})</div>
                                        ${client.emailsAffected.map(e => `
                                            <div style="background: white; border-radius: 6px; padding: 10px; margin-bottom: 6px;">
                                                <strong style="color: #333; font-size: 13px;">${Utils.escapeHtml(e.subject || 'Sans objet')}</strong>
                                                <div style="color: #666; font-size: 11px; margin-top: 4px;">De: ${Utils.escapeHtml(e.from || '')} → ${Utils.escapeHtml(e.affectedTo || '')}</div>
                                            </div>
                                        `).join('')}
                                    </div>
                                    ` : ''}

                                    ${client.aircallCalls.length > 0 ? `
                                    <!-- Appels avec résumés IA complets -->
                                    <div style="background: linear-gradient(135deg, #fff3e0 0%, #ffe0b2 100%); border-radius: 10px; padding: 15px; border-left: 4px solid #f57c00;">
                                        <div style="font-weight: 600; color: #e65100; margin-bottom: 10px; font-size: 14px;">📞 Appels téléphoniques (${client.aircallCalls.length})</div>
                                        ${client.aircallCalls.map(call => {
                                            const bgColor = call.type === 'sortant' ? '#fff8e1' : '#e8f5e9';
                                            const borderColor = call.type === 'sortant' ? '#ffb300' : '#66bb6a';
                                            const moodIcon = call.mood === 'Positif' ? '😊' : (call.mood === 'Négatif' ? '😟' : (call.mood === 'Neutre' ? '😐' : ''));
                                            return `
                                            <div style="background: ${bgColor}; border-radius: 8px; padding: 12px; margin-bottom: 8px; border-left: 3px solid ${borderColor};">
                                                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                                                    <span style="font-weight: 600; color: #333;">
                                                        ${call.type === 'sortant' ? '📤 Sortant' : '📥 Entrant'}
                                                        <span style="font-weight: normal; color: #666;">• ${call.duration || ''}</span>
                                                        ${moodIcon ? `<span style="margin-left: 8px;">${moodIcon}</span>` : ''}
                                                    </span>
                                                    <span style="color: #888; font-size: 11px;">${call.time || ''}</span>
                                                </div>
                                                ${call.summary ? `
                                                <div style="background: white; border-radius: 6px; padding: 10px; font-size: 12px; color: #555; line-height: 1.5;">
                                                    <div style="color: #ff8f00; font-size: 10px; font-weight: 600; margin-bottom: 4px;">💬 RÉSUMÉ IA</div>
                                                    ${Utils.escapeHtml(call.summary)}
                                                </div>
                                                ` : ''}
                                            </div>
                                            `;
                                        }).join('')}
                                    </div>
                                    ` : ''}

                                    ${client.tasksCompleted.length > 0 ? `
                                    <!-- Tâches terminées avec heure -->
                                    <div style="background: linear-gradient(135deg, #fff8e1 0%, #ffecb3 100%); border-radius: 10px; padding: 15px; border-left: 4px solid #ff8f00;">
                                        <div style="font-weight: 600; color: #e65100; margin-bottom: 10px; font-size: 14px;">✅ Tâches terminées (${client.tasksCompleted.length})</div>
                                        ${client.tasksCompleted.map((t, tIdx) => `
                                            <div style="background: white; border-radius: 6px; padding: 10px; margin-bottom: 6px;">
                                                <div style="display: flex; justify-content: space-between; align-items: flex-start;">
                                                    <strong style="color: #333; font-size: 13px;">${Utils.escapeHtml(t.title)}</strong>
                                                    ${t.closedTime ? `<span style="background: #ff8f00; color: white; padding: 2px 8px; border-radius: 10px; font-size: 10px;">⏰ ${t.closedTime}</span>` : ''}
                                                </div>
                                                ${t.content ? `
                                                    <div id="task_short_${cuid}_${tIdx}" style="color: #666; font-size: 12px; margin-top: 6px; line-height: 1.4;">${Utils.escapeHtml(Utils.truncate(t.content, 120))}</div>
                                                    ${t.content.length > 120 ? `
                                                        <div id="task_full_${cuid}_${tIdx}" style="display: none; color: #666; font-size: 12px; margin-top: 6px; line-height: 1.4; white-space: pre-wrap;">${Utils.escapeHtml(t.content)}</div>
                                                        <button onclick="var s=document.getElementById('task_short_${cuid}_${tIdx}');var f=document.getElementById('task_full_${cuid}_${tIdx}');if(f.style.display==='none'){f.style.display='block';s.style.display='none';this.textContent='▲ Réduire';}else{f.style.display='none';s.style.display='block';this.textContent='▼ Voir plus';}" style="background: #ff8f00; color: white; border: none; padding: 4px 10px; border-radius: 4px; font-size: 11px; cursor: pointer; margin-top: 6px;">▼ Voir plus</button>
                                                    ` : ''}
                                                ` : ''}
                                            </div>
                                        `).join('')}
                                    </div>
                                    ` : ''}

                                    ${client.tasksOverdue.length > 0 ? `
                                    <!-- Tâches en retard -->
                                    <div style="background: linear-gradient(135deg, #ffebee 0%, #ffcdd2 100%); border-radius: 10px; padding: 15px; border-left: 4px solid #d32f2f;">
                                        <div style="font-weight: 600; color: #c62828; margin-bottom: 10px; font-size: 14px;">⚠️ Tâches en retard (${client.tasksOverdue.length})</div>
                                        ${client.tasksOverdue.map(t => `
                                            <div style="background: white; border-radius: 6px; padding: 10px; margin-bottom: 6px;">
                                                <strong style="color: #333; font-size: 13px;">${Utils.escapeHtml(t.title)}</strong>
                                                <div style="color: #d32f2f; font-size: 11px; margin-top: 4px;">${t.daysOverdue}j de retard • → ${Utils.escapeHtml(t.assignedTo || 'N/A')}</div>
                                            </div>
                                        `).join('')}
                                    </div>
                                    ` : ''}

                                    ${(client.estimates.length > 0 || client.policies.length > 0 || client.claims.length > 0) ? `
                                    <!-- Documents -->
                                    <div style="background: linear-gradient(135deg, #f3e5f5 0%, #e1bee7 100%); border-radius: 10px; padding: 15px; border-left: 4px solid #7b1fa2;">
                                        <div style="font-weight: 600; color: #6a1b9a; margin-bottom: 10px; font-size: 14px;">📄 Documents (${client.estimates.length + client.policies.length + client.claims.length})</div>
                                        <div style="display: flex; flex-wrap: wrap; gap: 8px;">
                                            ${client.estimates.map(e => `<span style="background: white; color: #7b1fa2; padding: 6px 12px; border-radius: 6px; font-size: 12px;">📋 Devis ${e.entityId || ''}</span>`).join('')}
                                            ${client.policies.map(p => `<span style="background: white; color: #00796b; padding: 6px 12px; border-radius: 6px; font-size: 12px;">📄 Contrat ${p.entityId || ''}</span>`).join('')}
                                            ${client.claims.map(c => `<span style="background: white; color: #c62828; padding: 6px 12px; border-radius: 6px; font-size: 12px;">🚨 Sinistre ${c.entityId || ''}</span>`).join('')}
                                        </div>
                                    </div>
                                    ` : ''}

                                    ${client.logs.length > 0 ? `
                                    <!-- Modifications regroupées -->
                                    <div style="background: linear-gradient(135deg, #efebe9 0%, #d7ccc8 100%); border-radius: 10px; padding: 15px; border-left: 4px solid #5d4037;">
                                        <div style="font-weight: 600; color: #4e342e; margin-bottom: 8px; font-size: 14px;">📝 Modifications fiche (${client.logs.length})</div>
                                        <div style="display: flex; flex-wrap: wrap; gap: 8px;">
                                            ${(() => {
                                                const grouped = {};
                                                client.logs.forEach(l => {
                                                    const action = l.action || 'Modification';
                                                    const key = action.replace(/[✨✏️➕🗑️📝]/g, '').trim();
                                                    if (!grouped[key]) grouped[key] = 0;
                                                    grouped[key]++;
                                                });
                                                return Object.entries(grouped).map(([action, count]) => {
                                                    let icon = '📝';
                                                    let bg = '#f5f5f5';
                                                    if (action.toLowerCase().includes('création') || action.toLowerCase().includes('insert')) { icon = '✨'; bg = '#e8f5e9'; }
                                                    else if (action.toLowerCase().includes('modification') || action.toLowerCase().includes('update')) { icon = '✏️'; bg = '#fff3e0'; }
                                                    else if (action.toLowerCase().includes('suppression') || action.toLowerCase().includes('delete')) { icon = '🗑️'; bg = '#ffebee'; }
                                                    return `<span style="background: ${bg}; padding: 6px 12px; border-radius: 15px; font-size: 12px; font-weight: 500;">${icon} ${count} ${action.toLowerCase()}</span>`;
                                                }).join('');
                                            })()}
                                        </div>
                                    </div>
                                    ` : ''}
                                </div>
                            </div>
                            `;
                        }).join('') : '<div style="background: white; border-radius: 12px; padding: 50px; text-align: center; color: #666;">Aucun client concerné aujourd\'hui</div>'}

                        <!-- Footer -->
                        <div style="text-align: center; color: #999; padding-top: 15px; font-size: 12px;">
                            📊 Vue par Client - LTOA Modulr Script v4.7
                        </div>
                    </div>
                </div>
            `;

            // Afficher la vue
            document.body.insertAdjacentHTML('beforeend', clientViewHTML);

            // Event listeners
            document.getElementById('ltoa-close-client-view').addEventListener('click', () => {
                document.getElementById('ltoa-client-view-modal').remove();
            });

            } catch (error) {
                console.error('[LTOA-Report] Erreur Vue par Client:', error);
                alert('❌ Erreur lors de la génération de la vue par client. Consultez la console F12.');
            }
        },

        // ============================================
        // VUE CHRONOLOGIQUE
        // ============================================
        showChronoView() {
            try {
                const { emailsSent, emailsAffected, aircallCalls, tasksCompleted, logs, estimates, policies, claims, user, date } = this.data;

                // Collecter toutes les actions avec leur heure
                const allActions = [];

                // Parser l'heure d'une chaîne (format HH:MM ou HH:MM:SS)
                const parseTime = (timeStr) => {
                    if (!timeStr) return null;
                    // Chercher un pattern HH:MM ou HH:MM:SS
                    const match = timeStr.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
                    if (match) {
                        const hours = parseInt(match[1]);
                        const minutes = parseInt(match[2]);
                        const seconds = match[3] ? parseInt(match[3]) : 0;
                        return hours * 3600 + minutes * 60 + seconds;
                    }
                    return null;
                };

                // Extraire l'heure d'une date complète DD/MM/YYYY HH:MM
                const extractTime = (dateStr) => {
                    if (!dateStr) return null;
                    const match = dateStr.match(/(\d{1,2}):(\d{2})/);
                    if (match) {
                        return `${match[1].padStart(2, '0')}:${match[2]}`;
                    }
                    return null;
                };

                // Emails envoyés
                emailsSent.forEach(e => {
                    const time = extractTime(e.date);
                    if (time) {
                        allActions.push({
                            type: 'email_sent',
                            icon: '📤',
                            color: '#1976d2',
                            label: 'Email envoyé',
                            time: time,
                            timeSeconds: parseTime(time),
                            title: e.subject || 'Sans objet',
                            detail: `À: ${e.to || 'N/A'}`,
                            client: e.client || ''
                        });
                    }
                });

                // Emails affectés
                emailsAffected.forEach(e => {
                    const time = extractTime(e.date);
                    if (time) {
                        allActions.push({
                            type: 'email_affected',
                            icon: '📥',
                            color: '#388e3c',
                            label: 'Email affecté',
                            time: time,
                            timeSeconds: parseTime(time),
                            title: e.subject || 'Sans objet',
                            detail: `De: ${e.from || 'N/A'}`,
                            client: e.client || ''
                        });
                    }
                });

                // Appels Aircall
                (aircallCalls || []).forEach(c => {
                    const time = c.time;
                    if (time) {
                        allActions.push({
                            type: 'call',
                            icon: c.type === 'sortant' ? '📞↗' : '📞↙',
                            color: '#ff8f00',
                            label: c.type === 'sortant' ? 'Appel sortant' : 'Appel entrant',
                            time: time,
                            timeSeconds: parseTime(time),
                            title: c.contact || 'Inconnu',
                            detail: `Durée: ${c.duration || '0s'}${c.mood ? ' | ' + c.mood : ''}`,
                            client: c.contact || '',
                            summary: c.summary
                        });
                    }
                });

                // Tâches terminées (utiliser l'heure de clôture)
                tasksCompleted.forEach(t => {
                    const time = t.closedTime || extractTime(t.completedDate);
                    if (time) {
                        allActions.push({
                            type: 'task',
                            icon: '✅',
                            color: '#f57c00',
                            label: 'Tâche terminée',
                            time: time,
                            timeSeconds: parseTime(time),
                            title: t.title || 'Sans titre',
                            detail: `Client: ${t.client || 'N/A'}`,
                            client: t.client || '',
                            summary: t.content
                        });
                    }
                });

                // Logs/Actions diverses
                logs.forEach(l => {
                    const time = extractTime(l.date);
                    if (time) {
                        allActions.push({
                            type: 'log',
                            icon: '📝',
                            color: '#7b1fa2',
                            label: l.type || 'Action',
                            time: time,
                            timeSeconds: parseTime(time),
                            title: l.action || l.description || 'Action',
                            detail: `Client: ${l.client || 'N/A'}`,
                            client: l.client || ''
                        });
                    }
                });

                // Devis
                estimates.forEach(e => {
                    const time = extractTime(e.date);
                    if (time) {
                        allActions.push({
                            type: 'estimate',
                            icon: '📋',
                            color: '#0097a7',
                            label: 'Devis',
                            time: time,
                            timeSeconds: parseTime(time),
                            title: e.reference || 'Devis',
                            detail: `Client: ${e.client || 'N/A'}`,
                            client: e.client || ''
                        });
                    }
                });

                // Contrats
                policies.forEach(p => {
                    const time = extractTime(p.date);
                    if (time) {
                        allActions.push({
                            type: 'policy',
                            icon: '📄',
                            color: '#3f51b5',
                            label: 'Contrat',
                            time: time,
                            timeSeconds: parseTime(time),
                            title: p.reference || 'Contrat',
                            detail: `Client: ${p.client || 'N/A'}`,
                            client: p.client || ''
                        });
                    }
                });

                // Sinistres
                claims.forEach(c => {
                    const time = extractTime(c.date);
                    if (time) {
                        allActions.push({
                            type: 'claim',
                            icon: '🚨',
                            color: '#c2185b',
                            label: 'Sinistre',
                            time: time,
                            timeSeconds: parseTime(time),
                            title: c.reference || 'Sinistre',
                            detail: `Client: ${c.client || 'N/A'}`,
                            client: c.client || ''
                        });
                    }
                });

                // Trier par heure (croissant)
                allActions.sort((a, b) => {
                    if (a.timeSeconds === null) return 1;
                    if (b.timeSeconds === null) return -1;
                    return a.timeSeconds - b.timeSeconds;
                });

                // Calculer le temps écoulé entre chaque action
                const formatDuration = (seconds) => {
                    if (seconds < 60) return `${seconds}s`;
                    if (seconds < 3600) {
                        const mins = Math.floor(seconds / 60);
                        const secs = seconds % 60;
                        return secs > 0 ? `${mins}min ${secs}s` : `${mins}min`;
                    }
                    const hours = Math.floor(seconds / 3600);
                    const mins = Math.floor((seconds % 3600) / 60);
                    return mins > 0 ? `${hours}h ${mins}min` : `${hours}h`;
                };

                // Générer le HTML de la timeline
                let timelineHTML = '';
                for (let i = 0; i < allActions.length; i++) {
                    const action = allActions[i];
                    const prevAction = i > 0 ? allActions[i - 1] : null;

                    // Calculer le temps écoulé depuis l'action précédente
                    let elapsedHTML = '';
                    if (prevAction && action.timeSeconds !== null && prevAction.timeSeconds !== null) {
                        const elapsed = action.timeSeconds - prevAction.timeSeconds;
                        if (elapsed > 0) {
                            const elapsedFormatted = formatDuration(elapsed);
                            elapsedHTML = `
                                <div style="
                                    display: flex;
                                    align-items: center;
                                    padding: 8px 0;
                                    margin-left: 18px;
                                ">
                                    <div style="
                                        width: 2px;
                                        height: 30px;
                                        background: linear-gradient(to bottom, ${prevAction.color}, ${action.color});
                                        margin-right: 15px;
                                    "></div>
                                    <div style="
                                        background: #f5f5f5;
                                        padding: 4px 12px;
                                        border-radius: 12px;
                                        font-size: 11px;
                                        color: #666;
                                    ">
                                        ⏱️ ${elapsedFormatted}
                                    </div>
                                </div>
                            `;
                        }
                    }

                    timelineHTML += `
                        ${elapsedHTML}
                        <div style="
                            display: flex;
                            align-items: flex-start;
                            padding: 10px 0;
                        ">
                            <div style="
                                width: 38px;
                                height: 38px;
                                border-radius: 50%;
                                background: ${action.color};
                                display: flex;
                                align-items: center;
                                justify-content: center;
                                font-size: 16px;
                                flex-shrink: 0;
                                box-shadow: 0 2px 8px ${action.color}40;
                            ">${action.icon}</div>
                            <div style="
                                flex: 1;
                                margin-left: 15px;
                                background: white;
                                border: 1px solid #e0e0e0;
                                border-radius: 8px;
                                padding: 12px 15px;
                                box-shadow: 0 1px 3px rgba(0,0,0,0.05);
                            ">
                                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 5px;">
                                    <span style="
                                        background: ${action.color}20;
                                        color: ${action.color};
                                        padding: 2px 8px;
                                        border-radius: 4px;
                                        font-size: 11px;
                                        font-weight: bold;
                                    ">${action.label}</span>
                                    <span style="
                                        font-size: 13px;
                                        color: #333;
                                        font-weight: bold;
                                    ">🕐 ${action.time}</span>
                                </div>
                                <div style="font-size: 14px; font-weight: 600; color: #333; margin-bottom: 3px;">
                                    ${Utils.escapeHtml(action.title)}
                                </div>
                                <div style="font-size: 12px; color: #666;">
                                    ${Utils.escapeHtml(action.detail)}
                                </div>
                                ${action.summary ? `
                                    <div style="
                                        margin-top: 8px;
                                        padding: 8px;
                                        background: #f9f9f9;
                                        border-radius: 4px;
                                        font-size: 11px;
                                        color: #555;
                                        border-left: 3px solid ${action.color};
                                    ">
                                        ✨ ${Utils.escapeHtml(Utils.truncate(action.summary, 150))}
                                    </div>
                                ` : ''}
                            </div>
                        </div>
                    `;
                }

                // Calculer le temps total de travail
                let totalWorkTime = '';
                if (allActions.length >= 2) {
                    const first = allActions[0];
                    const last = allActions[allActions.length - 1];
                    if (first.timeSeconds !== null && last.timeSeconds !== null) {
                        const total = last.timeSeconds - first.timeSeconds;
                        totalWorkTime = `<div style="
                            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                            color: white;
                            padding: 15px 20px;
                            border-radius: 10px;
                            margin-bottom: 20px;
                            display: flex;
                            justify-content: space-between;
                            align-items: center;
                        ">
                            <div>
                                <div style="font-size: 12px; opacity: 0.9;">Plage horaire de travail</div>
                                <div style="font-size: 18px; font-weight: bold;">${first.time} → ${last.time}</div>
                            </div>
                            <div style="text-align: right;">
                                <div style="font-size: 12px; opacity: 0.9;">Durée totale</div>
                                <div style="font-size: 18px; font-weight: bold;">${formatDuration(total)}</div>
                            </div>
                        </div>`;
                    }
                }

                // Créer le modal
                const modalHTML = `
                    <div id="ltoa-chrono-modal" style="
                        position: fixed;
                        top: 0;
                        left: 0;
                        width: 100%;
                        height: 100%;
                        background: rgba(0,0,0,0.85);
                        z-index: 9999999;
                        overflow-y: auto;
                        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                    ">
                        <div style="
                            max-width: 800px;
                            margin: 20px auto;
                            background: #fafafa;
                            border-radius: 12px;
                            padding: 25px;
                            box-shadow: 0 10px 50px rgba(0,0,0,0.3);
                        ">
                            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;">
                                <div>
                                    <h2 style="margin: 0; color: #333; font-size: 20px;">🕐 Vue Chronologique</h2>
                                    <p style="margin: 5px 0 0 0; color: #666; font-size: 13px;">
                                        ${Utils.escapeHtml(user)} - ${date} - ${allActions.length} actions
                                    </p>
                                </div>
                                <button id="ltoa-close-chrono" style="
                                    background: #666;
                                    color: white;
                                    border: none;
                                    padding: 10px 20px;
                                    border-radius: 5px;
                                    cursor: pointer;
                                    font-size: 13px;
                                ">✕ Fermer</button>
                            </div>

                            ${totalWorkTime}

                            <div style="padding: 10px 0;">
                                ${allActions.length > 0 ? timelineHTML : `
                                    <p style="text-align: center; color: #666; padding: 40px;">
                                        Aucune action avec heure trouvée pour cette journée.
                                    </p>
                                `}
                            </div>
                        </div>
                    </div>
                `;

                document.body.insertAdjacentHTML('beforeend', modalHTML);

                document.getElementById('ltoa-close-chrono').addEventListener('click', () => {
                    document.getElementById('ltoa-chrono-modal').remove();
                });

                // Fermer avec Escape
                const escHandler = (e) => {
                    if (e.key === 'Escape') {
                        const modal = document.getElementById('ltoa-chrono-modal');
                        if (modal) {
                            modal.remove();
                            document.removeEventListener('keydown', escHandler);
                        }
                    }
                };
                document.addEventListener('keydown', escHandler);

            } catch (error) {
                console.error('[LTOA-Report] Erreur Vue Chronologique:', error);
                alert('❌ Erreur lors de la génération de la vue chronologique. Consultez la console F12.');
            }
        },

        exportHTML() {
            const { emailsSent, emailsAffected, aircallCalls, pendingEmailsCount, tasksCompleted, tasksOverdue, logs, estimates, policies, claims, user, date, notes, aircallStatus } = this.data;

            // Compteurs Aircall
            const aircallInbound = (aircallCalls || []).filter(c => c.type === 'entrant').length;
            const aircallOutbound = (aircallCalls || []).filter(c => c.type === 'sortant').length;
            const aircallAnswered = (aircallCalls || []).filter(c => c.answered !== false && !c.missedReason).length;
            const aircallTalkSeconds = (aircallCalls || []).reduce((sum, c) => sum + (Number(c.durationSeconds) || 0), 0);
            const aircallDurationText = `${Math.floor(aircallTalkSeconds / 3600)}h ${Math.floor((aircallTalkSeconds % 3600) / 60)}min`;
            const uniqueClients = new Set([
                ...emailsSent, ...emailsAffected, ...tasksCompleted,
                ...estimates, ...policies, ...claims, ...logs
            ].map(item => item.clientId || item.clientName || item.entityId)
              .filter(value => value && value !== 'N/A' && value !== 'Non associé')).size;
            const remainingWork = (pendingEmailsCount || 0) + tasksOverdue.length;

            // Générer un HTML statique complet (pas besoin de JS)
            const htmlContent = `<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Rapport d'Activité - ${Utils.escapeHtml(user)} - ${date}</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            background: #f5f5f5;
            padding: 20px;
            line-height: 1.6;
            color: #333;
        }
        .report-container {
            max-width: 1000px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        .header {
            background: linear-gradient(135deg, #c62828 0%, #8e0000 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        .header h1 { font-size: 28px; margin-bottom: 10px; }
        .header p { opacity: 0.9; font-size: 16px; }
        .summary {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 15px;
            padding: 25px;
            background: #fafafa;
        }
        .summary-card {
            text-align: center;
            padding: 20px 15px;
            border-radius: 12px;
            color: white;
        }
        .summary-card .number { font-size: 32px; font-weight: bold; }
        .summary-card .label { font-size: 12px; margin-top: 5px; }
        .bg-blue { background: linear-gradient(135deg, #1976d2, #0d47a1); }
        .bg-green { background: linear-gradient(135deg, #388e3c, #1b5e20); }
        .bg-orange { background: linear-gradient(135deg, #f57c00, #e65100); }
        .bg-red { background: linear-gradient(135deg, #d32f2f, #b71c1c); }
        .bg-cyan { background: linear-gradient(135deg, #0097a7, #006064); }
        .bg-indigo { background: linear-gradient(135deg, #3f51b5, #1a237e); }
        .bg-pink { background: linear-gradient(135deg, #c2185b, #880e4f); }
        .bg-purple { background: linear-gradient(135deg, #7b1fa2, #4a148c); }

        .section {
            padding: 25px;
            border-bottom: 1px solid #eee;
        }
        .section:last-child { border-bottom: none; }
        .section-title {
            font-size: 18px;
            font-weight: bold;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 3px solid;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .section-title.blue { color: #1976d2; border-color: #1976d2; }
        .section-title.green { color: #388e3c; border-color: #388e3c; }
        .section-title.orange { color: #f57c00; border-color: #f57c00; }
        .section-title.red { color: #d32f2f; border-color: #d32f2f; }
        .section-title.cyan { color: #0097a7; border-color: #0097a7; }
        .section-title.indigo { color: #3f51b5; border-color: #3f51b5; }
        .section-title.pink { color: #c2185b; border-color: #c2185b; }
        .section-title.purple { color: #7b1fa2; border-color: #7b1fa2; }

        table { width: 100%; border-collapse: collapse; font-size: 13px; }
        th { background: #f5f5f5; padding: 12px 10px; text-align: left; font-weight: 600; border-bottom: 2px solid #ddd; }
        td { padding: 10px; border-bottom: 1px solid #eee; vertical-align: top; }
        tr:hover { background: #fafafa; }

        .email-content, .task-content {
            background: #f8f9fa;
            padding: 10px;
            border-radius: 5px;
            margin-top: 8px;
            font-size: 12px;
            color: #555;
            border-left: 3px solid #ddd;
            max-height: 150px;
            overflow-y: auto;
        }

        .badge {
            display: inline-block;
            padding: 3px 8px;
            border-radius: 12px;
            font-size: 11px;
            font-weight: bold;
        }
        .badge-red { background: #ffebee; color: #c62828; }
        .badge-green { background: #e8f5e9; color: #2e7d32; }
        .badge-orange { background: #fff3e0; color: #e65100; }

        .footer {
            text-align: center;
            padding: 20px;
            background: #fafafa;
            color: #999;
            font-size: 12px;
        }

        a { color: #1976d2; text-decoration: none; }
        a:hover { text-decoration: underline; }

        .empty { color: #999; font-style: italic; text-align: center; padding: 30px; }

        @media print {
            body { background: white; padding: 0; }
            .report-container { box-shadow: none; }
            .section { page-break-inside: avoid; }
        }
        @media (max-width: 800px) {
            .summary { grid-template-columns: repeat(2, 1fr); }
        }
    </style>
</head>
<body>
    <!-- Boutons de navigation -->
    <div style="max-width: 1000px; margin: 0 auto 20px auto; display: flex; gap: 10px; justify-content: center; flex-wrap: wrap;">
        <button onclick="showView('categories')" id="btn-categories" class="nav-btn active" style="
            padding: 12px 25px;
            background: linear-gradient(135deg, #c62828, #b71c1c);
            color: white;
            border: none;
            border-radius: 8px;
            font-size: 14px;
            font-weight: bold;
            cursor: pointer;
            box-shadow: 0 3px 10px rgba(0,0,0,0.2);
        ">📊 Vue Catégories</button>
        <button onclick="showView('chrono')" id="btn-chrono" class="nav-btn" style="
            padding: 12px 25px;
            background: linear-gradient(135deg, #9c27b0, #7b1fa2);
            color: white;
            border: none;
            border-radius: 8px;
            font-size: 14px;
            font-weight: bold;
            cursor: pointer;
            box-shadow: 0 3px 10px rgba(0,0,0,0.2);
        ">🕐 Vue Chronologique</button>
        <button onclick="showView('client')" id="btn-client" class="nav-btn" style="
            padding: 12px 25px;
            background: linear-gradient(135deg, #1565c0, #0d47a1);
            color: white;
            border: none;
            border-radius: 8px;
            font-size: 14px;
            font-weight: bold;
            cursor: pointer;
            box-shadow: 0 3px 10px rgba(0,0,0,0.2);
        ">👤 Vue par Client</button>
        <button onclick="window.print()" style="
            padding: 12px 25px;
            background: linear-gradient(135deg, #666, #444);
            color: white;
            border: none;
            border-radius: 8px;
            font-size: 14px;
            font-weight: bold;
            cursor: pointer;
            box-shadow: 0 3px 10px rgba(0,0,0,0.2);
        ">🖨️ Imprimer</button>
    </div>

    <div id="view-categories" class="report-container">
        <!-- Header -->
        <div class="header">
            <h1>📊 Rapport d'Activité Quotidien</h1>
            <p><strong>${Utils.escapeHtml(user)}</strong> - ${date}</p>
        </div>

        ${aircallStatus?.state === 'error' ? `
        <div style="margin:0 25px 20px;padding:12px 16px;border-radius:8px;background:#ffebee;color:#b71c1c;font-size:12px;">
            ⚠️ ${Utils.escapeHtml(aircallStatus.message || 'Échec de la collecte Aircall')}
        </div>` : ''}

        <!-- Alerte emails en attente -->
        ${(pendingEmailsCount || 0) > 0 ? `
        <div style="background: linear-gradient(135deg, #ffcccb 0%, #ff6b6b 100%); padding: 15px 20px; border-radius: 10px; margin-bottom: 20px; display: flex; align-items: center; justify-content: space-between;">
            <div style="display: flex; align-items: center; gap: 15px;">
                <div style="font-size: 32px;">📬</div>
                <div>
                    <div style="font-size: 14px; color: #7f0000; font-weight: bold;">Emails assignés à ${Utils.escapeHtml(user)}</div>
                </div>
            </div>
            <div style="font-size: 36px; font-weight: bold; color: #b71c1c;">${pendingEmailsCount || 0}</div>
        </div>
        ` : ''}

        <!-- Summary -->
        <div class="summary">
            <div class="summary-card bg-blue">
                <div class="number">${emailsSent.length}</div>
                <div class="label">📤 Emails Envoyés</div>
            </div>
            <div class="summary-card bg-green">
                <div class="number">${emailsAffected.length}</div>
                <div class="label">📥 Emails Affectés</div>
            </div>
            <div class="summary-card" style="background: linear-gradient(135deg, #ff8f00, #e65100);">
                <div class="number">${(aircallCalls || []).length}</div>
                <div class="label">📞 Appels (${aircallInbound}↓ ${aircallOutbound}↑)</div>
            </div>
            <div class="summary-card bg-orange">
                <div class="number">${tasksCompleted.length}</div>
                <div class="label">✅ Tâches Terminées</div>
            </div>
            <div class="summary-card bg-red">
                <div class="number">${tasksOverdue.length}</div>
                <div class="label">⚠️ Tâches en Retard</div>
            </div>
            <div class="summary-card bg-cyan">
                <div class="number">${estimates.length}</div>
                <div class="label">📋 Devis</div>
            </div>
            <div class="summary-card bg-indigo">
                <div class="number">${policies.length}</div>
                <div class="label">📄 Contrats</div>
            </div>
            <div class="summary-card bg-pink">
                <div class="number">${claims.length}</div>
                <div class="label">🚨 Sinistres</div>
            </div>
            <div class="summary-card bg-purple">
                <div class="number">${logs.length}</div>
                <div class="label">📝 Autres Actions</div>
            </div>
        </div>

        <div class="section" style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px;">
            <div style="padding:14px;background:#e3f2fd;border-radius:8px;text-align:center;"><strong style="font-size:22px;color:#1565c0;">${uniqueClients}</strong><br><small>Dossiers/clients traités</small></div>
            <div style="padding:14px;background:#e8f5e9;border-radius:8px;text-align:center;"><strong style="font-size:22px;color:#2e7d32;">${estimates.length + policies.length}</strong><br><small>Actions de production</small></div>
            <div style="padding:14px;background:#fff3e0;border-radius:8px;text-align:center;"><strong style="font-size:22px;color:#ef6c00;">${tasksCompleted.length}</strong><br><small>Suivis finalisés</small></div>
            <div style="padding:14px;background:${remainingWork ? '#ffebee' : '#e8f5e9'};border-radius:8px;text-align:center;"><strong style="font-size:22px;color:${remainingWork ? '#c62828' : '#2e7d32'};">${remainingWork}</strong><br><small>Éléments restant à traiter</small></div>
        </div>

        <!-- Emails Envoyés -->
        ${emailsSent.length > 0 ? `
        <div class="section">
            <h2 class="section-title blue">📤 Emails Envoyés (${emailsSent.length})</h2>
            <table>
                <tr>
                    <th style="width: 100px;">Date</th>
                    <th style="width: 200px;">Destinataire</th>
                    <th>Objet</th>
                </tr>
                ${emailsSent.map(e => `
                <tr>
                    <td>${e.date || ''} ${e.time || ''}</td>
                    <td>
                        ${e.clientName ? `<strong>${Utils.escapeHtml(e.clientName)}</strong><br>` : ''}
                        <span style="color: #666;">${Utils.escapeHtml(e.toEmail || e.to || '')}</span>
                    </td>
                    <td>
                        <strong>${Utils.escapeHtml(e.subject || 'Sans objet')}</strong>
                        ${e.body ? `<div class="email-content">${Utils.escapeHtml(e.body)}</div>` : ''}
                    </td>
                </tr>
                `).join('')}
            </table>
        </div>
        ` : ''}

        <!-- Emails Affectés -->
        ${emailsAffected.length > 0 ? `
        <div class="section">
            <h2 class="section-title green">📥 Emails Affectés (${emailsAffected.length})</h2>
            <table>
                <tr>
                    <th style="width: 100px;">Date</th>
                    <th style="width: 180px;">Expéditeur</th>
                    <th>Objet</th>
                    <th style="width: 150px;">Affecté à</th>
                </tr>
                ${emailsAffected.map(e => `
                <tr>
                    <td>${e.date || ''} ${e.time || ''}</td>
                    <td>${Utils.escapeHtml(e.from || e.fromEmail || '')}</td>
                    <td><strong>${Utils.escapeHtml(e.subject || 'Sans objet')}</strong></td>
                    <td>${Utils.escapeHtml(e.affectedTo || '')}</td>
                </tr>
                `).join('')}
            </table>
        </div>
        ` : ''}

        <!-- Appels Téléphoniques Aircall -->
        ${(aircallCalls || []).length > 0 ? `
        <div class="section">
            <h2 class="section-title" style="color: #ff8f00; border-color: #ff8f00;">📞 Appels Téléphoniques (${(aircallCalls || []).length}) - ${aircallInbound} entrants / ${aircallOutbound} sortants</h2>
            <table>
                <tr>
                    <th style="width: 80px;">Heure</th>
                    <th style="width: 80px;">Type</th>
                    <th style="width: 150px;">Contact</th>
                    <th style="width: 80px;">Durée</th>
                    <th style="width: 80px;">Humeur</th>
                    <th>Analyse de l’appel</th>
                </tr>
                ${(aircallCalls || []).map(c => `
                <tr>
                    <td>${c.time || ''}</td>
                    <td>${c.type === 'sortant' ? '📤 Sortant' : '📥 Entrant'}</td>
                    <td><strong>${Utils.escapeHtml(c.contact || 'Inconnu')}</strong></td>
                    <td>${Utils.escapeHtml(c.duration || '0s')}</td>
                    <td>${c.mood === 'Positif' ? '😊 Positif' : (c.mood === 'Négatif' ? '😟 Négatif' : (c.mood === 'Neutre' ? '😐 Neutre' : '-'))}</td>
                    <td style="font-size: 12px;">
                        ${c.summary ? `<strong>Résumé :</strong> ${Utils.escapeHtml(c.summary)}` : '<span style="color:#999;">Résumé indisponible</span>'}
                        ${c.topics?.length ? `<br><strong>Sujets clés :</strong> ${c.topics.map(topic => Utils.escapeHtml(topic)).join(' · ')}` : ''}
                        ${c.actionItems?.length ? `<br><strong>Actions :</strong><ul>${c.actionItems.map(action => `<li>${Utils.escapeHtml(action)}</li>`).join('')}</ul>` : ''}
                        ${c.transcript ? `<details><summary>Voir la transcription</summary><div style="white-space:pre-wrap;font-size:11px;">${Utils.escapeHtml(c.transcript)}</div></details>` : ''}
                    </td>
                </tr>
                `).join('')}
            </table>
        </div>
        ` : ''}

        <!-- Tâches Terminées -->
        ${tasksCompleted.length > 0 ? `
        <div class="section">
            <h2 class="section-title orange">✅ Tâches Terminées (${tasksCompleted.length})</h2>
            <table>
                <tr>
                    <th style="width: 70px;">Heure</th>
                    <th style="width: 220px;">Tâche</th>
                    <th style="width: 150px;">Client</th>
                    <th>Contenu</th>
                </tr>
                ${tasksCompleted.map(t => `
                <tr>
                    <td style="text-align: center; font-weight: bold; color: #f57c00;">${t.closedTime || '-'}</td>
                    <td><strong>${Utils.escapeHtml(t.title || '')}</strong></td>
                    <td>
                        ${t.clientId ? `<a href="https://courtage.modulr.fr/fr/scripts/clients/clients_card.php?id=${t.clientId}" target="_blank">` : ''}
                        ${Utils.escapeHtml(t.clientName || t.client || 'Non associé')}
                        ${t.clientId ? '</a>' : ''}
                    </td>
                    <td>${t.content ? `<div class="task-content">${Utils.escapeHtml(t.content)}</div>` : '<span style="color:#999;">-</span>'}</td>
                </tr>
                `).join('')}
            </table>
        </div>
        ` : ''}

        <!-- Tâches en Retard -->
        ${tasksOverdue.length > 0 ? `
        <div class="section">
            <h2 class="section-title red">⚠️ Tâches en Retard (${tasksOverdue.length})</h2>
            <table>
                <tr>
                    <th style="width: 250px;">Tâche</th>
                    <th style="width: 150px;">Client</th>
                    <th>Contenu</th>
                    <th style="width: 100px;">Retard</th>
                </tr>
                ${tasksOverdue.map(t => `
                <tr>
                    <td><strong>${Utils.escapeHtml(t.title || '')}</strong></td>
                    <td>${Utils.escapeHtml(t.clientName || t.client || 'Non associé')}</td>
                    <td>${t.content ? `<div class="task-content">${Utils.escapeHtml(t.content)}</div>` : '<span style="color:#999;">-</span>'}</td>
                    <td><span class="badge badge-red">${t.daysOverdue || '?'}j de retard</span></td>
                </tr>
                `).join('')}
            </table>
        </div>
        ` : ''}

        <!-- Devis -->
        ${estimates.length > 0 ? `
        <div class="section">
            <h2 class="section-title cyan">📋 Devis (${estimates.length})</h2>
            <table>
                <tr>
                    <th style="width: 100px;">Action</th>
                    <th style="width: 200px;">Devis</th>
                    <th>Modifications</th>
                    <th style="width: 120px;">Date</th>
                </tr>
                ${estimates.map(e => `
                <tr>
                    <td>${e.action || ''}</td>
                    <td>${Utils.escapeHtml(e.entityName || '')}</td>
                    <td>${e.changes && e.changes.length > 0 ? e.changes.map(c => `<strong>${c.field}</strong>: ${c.oldValue} → ${c.newValue}`).join('<br>') : '<span style="color:#999;">Création</span>'}</td>
                    <td>${e.date || ''}</td>
                </tr>
                `).join('')}
            </table>
        </div>
        ` : ''}

        <!-- Contrats -->
        ${policies.length > 0 ? `
        <div class="section">
            <h2 class="section-title indigo">📄 Contrats (${policies.length})</h2>
            <table>
                <tr>
                    <th style="width: 100px;">Action</th>
                    <th style="width: 200px;">Contrat</th>
                    <th>Modifications</th>
                    <th style="width: 120px;">Date</th>
                </tr>
                ${policies.map(p => `
                <tr>
                    <td>${p.action || ''}</td>
                    <td>${Utils.escapeHtml(p.entityName || '')}</td>
                    <td>${p.changes && p.changes.length > 0 ? p.changes.map(c => `<strong>${c.field}</strong>: ${c.oldValue} → ${c.newValue}`).join('<br>') : '<span style="color:#999;">Création</span>'}</td>
                    <td>${p.date || ''}</td>
                </tr>
                `).join('')}
            </table>
        </div>
        ` : ''}

        <!-- Sinistres -->
        ${claims.length > 0 ? `
        <div class="section">
            <h2 class="section-title pink">🚨 Sinistres (${claims.length})</h2>
            <table>
                <tr>
                    <th style="width: 100px;">Action</th>
                    <th style="width: 200px;">Sinistre</th>
                    <th>Modifications</th>
                    <th style="width: 120px;">Date</th>
                </tr>
                ${claims.map(c => `
                <tr>
                    <td>${c.action || ''}</td>
                    <td>${Utils.escapeHtml(c.entityName || '')}</td>
                    <td>${c.changes && c.changes.length > 0 ? c.changes.map(ch => `<strong>${ch.field}</strong>: ${ch.oldValue} → ${ch.newValue}`).join('<br>') : '<span style="color:#999;">Création</span>'}</td>
                    <td>${c.date || ''}</td>
                </tr>
                `).join('')}
            </table>
        </div>
        ` : ''}

        <!-- Autres Actions -->
        ${logs.length > 0 ? `
        <div class="section">
            <h2 class="section-title purple">📝 Autres Actions (${logs.length})</h2>
            <table>
                <tr>
                    <th style="width: 100px;">Action</th>
                    <th style="width: 100px;">Type</th>
                    <th style="width: 180px;">Fiche</th>
                    <th>Modifications</th>
                    <th style="width: 120px;">Date</th>
                </tr>
                ${logs.map(l => `
                <tr>
                    <td>${l.action || ''}</td>
                    <td>${Utils.escapeHtml(l.table || '')}</td>
                    <td>${Utils.escapeHtml(l.entityName || '')}</td>
                    <td>${l.changes && l.changes.length > 0 ? l.changes.slice(0, 5).map(c => `<strong>${c.field}</strong>: ${c.oldValue} → ${c.newValue}`).join('<br>') + (l.changes.length > 5 ? '<br><em>+ ' + (l.changes.length - 5) + ' autres...</em>' : '') : '<span style="color:#999;">-</span>'}</td>
                    <td>${l.date || ''}</td>
                </tr>
                `).join('')}
            </table>
        </div>
        ` : ''}


        ${notes ? `
        <div class="section" style="background:#fffde7;">
            <h2 class="section-title" style="color:#6d4c00;border-color:#f9a825;">🗒️ Notes et précisions</h2>
            <div style="white-space:pre-wrap;">${Utils.escapeHtml(notes)}</div>
        </div>` : ''}

        <!-- Footer -->
        <div class="footer">
            <p>Rapport généré le ${new Date().toLocaleString('fr-FR')} par LTOA Modulr Script v4.7</p>
        </div>
    </div>

    <!-- Vue Chronologique (cachée par défaut) -->
    <div id="view-chrono" class="report-container" style="display: none;">
        <div class="header" style="background: linear-gradient(135deg, #9c27b0, #7b1fa2);">
            <h1>🕐 Vue Chronologique</h1>
            <p><strong>${Utils.escapeHtml(user)}</strong> - ${date}</p>
        </div>
        <div style="padding: 20px;">
            ${this.generateChronoViewHTML()}
        </div>
        <div class="footer">
            <p>Rapport généré le ${new Date().toLocaleString('fr-FR')} par LTOA Modulr Script v4.7</p>
        </div>
    </div>

    <!-- Vue par Client (cachée par défaut) -->
    <div id="view-client" class="report-container" style="display: none;">
        <div class="header" style="background: linear-gradient(135deg, #1565c0, #0d47a1);">
            <h1>👤 Vue par Client</h1>
            <p>Toutes les actions groupées par client</p>
        </div>
        <div style="padding: 20px;">
            ${this.generateClientViewHTML()}
        </div>
        <div class="footer">
            <p>Rapport généré le ${new Date().toLocaleString('fr-FR')} par LTOA Modulr Script v4.7</p>
        </div>
    </div>

    <script>
        function showView(viewName) {
            // Cacher toutes les vues
            document.getElementById('view-categories').style.display = 'none';
            document.getElementById('view-chrono').style.display = 'none';
            document.getElementById('view-client').style.display = 'none';

            // Afficher la vue demandée
            document.getElementById('view-' + viewName).style.display = 'block';

            // Mettre à jour les styles des boutons
            document.getElementById('btn-categories').style.opacity = '0.6';
            document.getElementById('btn-chrono').style.opacity = '0.6';
            document.getElementById('btn-client').style.opacity = '0.6';
            document.getElementById('btn-' + viewName).style.opacity = '1';
        }

        // Fonction pour déplier/replier les détails client
        function toggleClient(id) {
            const el = document.getElementById(id);
            const icon = document.getElementById('icon-' + id);
            if (el.style.display === 'none') {
                el.style.display = 'block';
                icon.textContent = '▼';
            } else {
                el.style.display = 'none';
                icon.textContent = '▶';
            }
        }
    </script>
</body>
</html>`;

            // Créer le blob et télécharger
            const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `Rapport_${user.replace(/\s+/g, '_')}_${date.replace(/\//g, '-')}.html`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);

            alert(`✅ HTML exporté: ${a.download}\n\nVous pouvez l'ouvrir dans n'importe quel navigateur, l'imprimer ou le partager par email !`);
        },

        // Générer le HTML de la vue chronologique pour l'export
        generateChronoViewHTML() {
            const { emailsSent, emailsAffected, aircallCalls, tasksCompleted, tasksOverdue, estimates, policies, claims, logs, user, date } = this.data;

            // Fonctions utilitaires
            const parseTime = (timeStr) => {
                if (!timeStr) return null;
                const parts = timeStr.split(':');
                if (parts.length >= 2) {
                    const hours = parseInt(parts[0], 10);
                    const minutes = parseInt(parts[1], 10);
                    const seconds = parts[2] ? parseInt(parts[2], 10) : 0;
                    return hours * 3600 + minutes * 60 + seconds;
                }
                return null;
            };

            const extractTime = (dateStr) => {
                if (!dateStr) return null;
                const match = dateStr.match(/(\d{1,2}):(\d{2})/);
                if (match) {
                    return `${match[1].padStart(2, '0')}:${match[2]}`;
                }
                return null;
            };

            const formatDuration = (seconds) => {
                const h = Math.floor(seconds / 3600);
                const m = Math.floor((seconds % 3600) / 60);
                if (h > 0) return `${h}h ${m}min`;
                return `${m}min`;
            };

            // Collecter toutes les actions
            const allActions = [];

            // Emails envoyés
            emailsSent.forEach(e => {
                const time = e.time || extractTime(e.date);
                if (time) {
                    allActions.push({
                        type: 'email_sent',
                        icon: '📤',
                        color: '#1976d2',
                        label: 'Email envoyé',
                        time: time,
                        timeSeconds: parseTime(time),
                        title: e.subject || 'Sans objet',
                        detail: `À: ${e.to || e.toEmail || e.clientName || 'N/A'}`,
                        summary: e.body
                    });
                }
            });

            // Emails affectés
            emailsAffected.forEach(e => {
                const time = e.time || extractTime(e.date);
                if (time) {
                    allActions.push({
                        type: 'email_affected',
                        icon: '📥',
                        color: '#388e3c',
                        label: 'Email affecté',
                        time: time,
                        timeSeconds: parseTime(time),
                        title: e.subject || 'Sans objet',
                        detail: `De: ${e.from || 'N/A'} → ${e.affectedTo || 'N/A'}`
                    });
                }
            });

            // Appels Aircall
            (aircallCalls || []).forEach(c => {
                const time = c.time;
                if (time) {
                    allActions.push({
                        type: 'call',
                        icon: c.type === 'sortant' ? '📞↗' : '📞↙',
                        color: '#ff8f00',
                        label: c.type === 'sortant' ? 'Appel sortant' : 'Appel entrant',
                        time: time,
                        timeSeconds: parseTime(time),
                        title: c.contact || 'Inconnu',
                        detail: `Durée: ${c.duration || '0s'}${c.mood ? ' | ' + c.mood : ''}`,
                        summary: c.summary
                    });
                }
            });

            // Tâches terminées
            tasksCompleted.forEach(t => {
                const time = t.closedTime || extractTime(t.completedDate);
                if (time) {
                    allActions.push({
                        type: 'task',
                        icon: '✅',
                        color: '#f57c00',
                        label: 'Tâche terminée',
                        time: time,
                        timeSeconds: parseTime(time),
                        title: t.title || 'Sans titre',
                        detail: `Client: ${t.client || 'N/A'}`,
                        summary: t.content
                    });
                }
            });

            // Trier par heure
            allActions.sort((a, b) => {
                if (a.timeSeconds === null) return 1;
                if (b.timeSeconds === null) return -1;
                return a.timeSeconds - b.timeSeconds;
            });

            if (allActions.length === 0) {
                return '<p style="text-align: center; color: #999; padding: 40px;">Aucune activité avec heure à afficher</p>';
            }

            // Générer le bandeau temps de travail
            let totalWorkTimeHTML = '';
            if (allActions.length >= 2) {
                const first = allActions[0];
                const last = allActions[allActions.length - 1];
                if (first.timeSeconds !== null && last.timeSeconds !== null) {
                    const total = last.timeSeconds - first.timeSeconds;
                    totalWorkTimeHTML = `
                    <div style="
                        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                        color: white;
                        padding: 15px 20px;
                        border-radius: 10px;
                        margin-bottom: 20px;
                        display: flex;
                        justify-content: space-between;
                        align-items: center;
                    ">
                        <div>
                            <div style="font-size: 12px; opacity: 0.9;">Plage horaire de travail</div>
                            <div style="font-size: 18px; font-weight: bold;">${first.time} → ${last.time}</div>
                        </div>
                        <div style="text-align: right;">
                            <div style="font-size: 12px; opacity: 0.9;">Durée totale</div>
                            <div style="font-size: 18px; font-weight: bold;">${formatDuration(total)}</div>
                        </div>
                    </div>`;
                }
            }

            // Générer la timeline
            let timelineHTML = totalWorkTimeHTML;
            let prevAction = null;

            for (const action of allActions) {
                // Calculer temps écoulé depuis l'action précédente
                let elapsedHTML = '';
                if (prevAction && prevAction.timeSeconds !== null && action.timeSeconds !== null) {
                    const elapsed = action.timeSeconds - prevAction.timeSeconds;
                    if (elapsed > 60) {
                        elapsedHTML = `
                            <div style="text-align: center; padding: 8px 0; color: #999; font-size: 11px;">
                                ⏱️ +${formatDuration(elapsed)}
                            </div>`;
                    }
                }
                prevAction = action;

                timelineHTML += `
                    ${elapsedHTML}
                    <div style="display: flex; align-items: flex-start; padding: 10px 0;">
                        <div style="
                            width: 38px;
                            height: 38px;
                            border-radius: 50%;
                            background: ${action.color};
                            display: flex;
                            align-items: center;
                            justify-content: center;
                            font-size: 16px;
                            flex-shrink: 0;
                            box-shadow: 0 2px 8px ${action.color}40;
                        ">${action.icon}</div>
                        <div style="
                            flex: 1;
                            margin-left: 15px;
                            background: white;
                            border: 1px solid #e0e0e0;
                            border-radius: 8px;
                            padding: 12px 15px;
                            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
                        ">
                            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 5px;">
                                <span style="
                                    background: ${action.color}20;
                                    color: ${action.color};
                                    padding: 2px 8px;
                                    border-radius: 4px;
                                    font-size: 11px;
                                    font-weight: bold;
                                ">${action.label}</span>
                                <span style="font-size: 13px; color: #333; font-weight: bold;">🕐 ${action.time}</span>
                            </div>
                            <div style="font-size: 14px; font-weight: 600; color: #333; margin-bottom: 3px;">
                                ${Utils.escapeHtml(action.title)}
                            </div>
                            <div style="font-size: 12px; color: #666;">
                                ${Utils.escapeHtml(action.detail)}
                            </div>
                            ${action.summary ? `
                                <div style="
                                    margin-top: 8px;
                                    padding: 8px;
                                    background: #f9f9f9;
                                    border-radius: 4px;
                                    font-size: 11px;
                                    color: #555;
                                    border-left: 3px solid ${action.color};
                                ">
                                    ✨ ${Utils.escapeHtml(Utils.truncate(action.summary, 200))}
                                </div>
                            ` : ''}
                        </div>
                    </div>`;
            }

            return `<div style="background: #fafafa; padding: 15px; border-radius: 10px;">${timelineHTML}</div>`;
        },

        // Générer le HTML de la vue par client pour l'export
        generateClientViewHTML() {
            const { emailsSent, emailsAffected, aircallCalls, tasksCompleted, tasksOverdue, estimates, policies, claims, logs, clientIndex } = this.data;

            // Grouper par client (même logique que la modale)
            const clientsMap = new Map();

            const addToClient = (clientKey, clientName, clientId, clientEmail, category, item) => {
                if (!clientKey || clientKey === 'unknown' || clientKey === 'non_associé') {
                    clientKey = 'non_associé';
                    clientName = 'Non associé / Non résolu';
                }

                if (!clientsMap.has(clientKey)) {
                    clientsMap.set(clientKey, {
                        name: clientName || clientKey,
                        id: clientId || null,
                        email: clientEmail || null,
                        emailsSent: [],
                        emailsAffected: [],
                        aircallCalls: [],
                        tasksCompleted: [],
                        tasksOverdue: [],
                        estimates: [],
                        policies: [],
                        claims: [],
                        logs: []
                    });
                }

                const client = clientsMap.get(clientKey);
                if (clientName && clientName !== client.name) client.name = clientName;
                if (clientId && !client.id) client.id = clientId;
                if (clientEmail && !client.email) client.email = clientEmail;

                if (client[category]) {
                    client[category].push(item);
                }
            };

            // Grouper les données
            emailsSent.forEach(e => {
                const key = e.clientId || e.toEmail?.toLowerCase() || 'unknown';
                addToClient(key, e.clientName, e.clientId, e.toEmail, 'emailsSent', e);
            });

            emailsAffected.forEach(e => {
                const key = e.clientId || e.affectedTo?.toLowerCase() || 'unknown';
                addToClient(key, e.clientName || e.affectedTo, e.clientId, e.clientEmail, 'emailsAffected', e);
            });

            (aircallCalls || []).forEach(c => {
                const key = c.clientId || c.contact?.toLowerCase() || 'unknown';
                addToClient(key, c.contact, c.clientId, null, 'aircallCalls', c);
            });

            tasksCompleted.forEach(t => {
                const key = t.clientId || t.client?.toLowerCase() || 'unknown';
                addToClient(key, t.clientName || t.client, t.clientId, t.clientEmail, 'tasksCompleted', t);
            });

            tasksOverdue.forEach(t => {
                const key = t.clientId || t.client?.toLowerCase() || 'unknown';
                addToClient(key, t.clientName || t.client, t.clientId, t.clientEmail, 'tasksOverdue', t);
            });

            estimates.forEach(e => {
                const key = e.clientId || 'unknown';
                addToClient(key, e.clientName, e.clientId, e.clientEmail, 'estimates', e);
            });

            policies.forEach(p => {
                const key = p.clientId || 'unknown';
                addToClient(key, p.clientName, p.clientId, p.clientEmail, 'policies', p);
            });

            claims.forEach(c => {
                const key = c.clientId || 'unknown';
                addToClient(key, c.clientName, c.clientId, c.clientEmail, 'claims', c);
            });

            // Trier
            const sortedClients = Array.from(clientsMap.values()).sort((a, b) => {
                if (a.name === 'Non associé / Non résolu') return 1;
                if (b.name === 'Non associé / Non résolu') return -1;
                if (a.id && !b.id) return -1;
                if (!a.id && b.id) return 1;
                return (a.name || '').localeCompare(b.name || '');
            }).filter(c =>
                c.emailsSent.length + c.emailsAffected.length + c.aircallCalls.length +
                c.tasksCompleted.length + c.tasksOverdue.length + c.estimates.length +
                c.policies.length + c.claims.length > 0
            );

            if (sortedClients.length === 0) {
                return '<p style="text-align: center; color: #999; padding: 40px;">Aucun client trouvé</p>';
            }

            // Générer le HTML identique à la modale
            let html = '<div style="background: #f5f5f5; padding: 15px; border-radius: 10px;">';

            sortedClients.forEach((client, clientIdx) => {
                const clientLink = client.id ?
                    `https://courtage.modulr.fr/fr/scripts/clients/clients_card.php?id=${client.id}` : '#';
                const cuid = 'exp_' + clientIdx;

                html += `
                <div style="background: white; border-radius: 12px; margin-bottom: 20px; overflow: hidden; box-shadow: 0 2px 10px rgba(0,0,0,0.08);">
                    <!-- En-tête client avec gradient -->
                    <div style="background: linear-gradient(135deg, #1565c0 0%, #0d47a1 100%); color: white; padding: 18px 22px;">
                        <div style="display: flex; justify-content: space-between; align-items: flex-start;">
                            <div>
                                <a href="${clientLink}" target="_blank" style="color: white; text-decoration: none; font-size: 18px; font-weight: 600;">
                                    👤 ${Utils.escapeHtml(client.name)}
                                </a>
                                ${client.id ? `<span style="background: rgba(255,255,255,0.2); padding: 3px 10px; border-radius: 12px; font-size: 11px; margin-left: 10px;">N° ${client.id}</span>` : ''}
                                ${client.email ? `<div style="opacity: 0.8; font-size: 12px; margin-top: 5px;">📧 ${Utils.escapeHtml(client.email)}</div>` : ''}
                            </div>
                            <div style="display: flex; gap: 6px; flex-wrap: wrap;">
                                ${client.emailsSent.length > 0 ? `<span style="background: #2196f3; padding: 4px 12px; border-radius: 15px; font-size: 11px; font-weight: 500;">📤 ${client.emailsSent.length}</span>` : ''}
                                ${client.emailsAffected.length > 0 ? `<span style="background: #4caf50; padding: 4px 12px; border-radius: 15px; font-size: 11px; font-weight: 500;">📥 ${client.emailsAffected.length}</span>` : ''}
                                ${client.aircallCalls.length > 0 ? `<span style="background: #ff9800; padding: 4px 12px; border-radius: 15px; font-size: 11px; font-weight: 500;">📞 ${client.aircallCalls.length}</span>` : ''}
                                ${client.tasksCompleted.length > 0 ? `<span style="background: #ff5722; padding: 4px 12px; border-radius: 15px; font-size: 11px; font-weight: 500;">✅ ${client.tasksCompleted.length}</span>` : ''}
                            </div>
                        </div>
                    </div>

                    <!-- Contenu avec cartes colorées -->
                    <div style="padding: 18px; display: grid; gap: 12px;">

                        ${client.emailsSent.length > 0 ? `
                        <div style="background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%); border-radius: 10px; padding: 15px; border-left: 4px solid #1976d2;">
                            <div style="font-weight: 600; color: #1565c0; margin-bottom: 10px; font-size: 14px;">📤 Emails envoyés (${client.emailsSent.length})</div>
                            ${client.emailsSent.map((e, eIdx) => `
                                <div style="background: white; border-radius: 6px; padding: 10px; margin-bottom: 6px;">
                                    <div style="display: flex; justify-content: space-between;">
                                        <strong style="color: #333; font-size: 13px;">${Utils.escapeHtml(e.subject || 'Sans objet')}</strong>
                                        <span style="color: #1976d2; font-size: 11px; font-weight: 500;">${e.time || ''}</span>
                                    </div>
                                    ${e.body ? `
                                        <div id="email_short_${cuid}_${eIdx}" style="color: #666; font-size: 12px; margin-top: 6px; line-height: 1.4;">${Utils.escapeHtml(Utils.truncate(e.body, 150))}</div>
                                        ${e.body.length > 150 ? `
                                            <div id="email_full_${cuid}_${eIdx}" style="display: none; color: #666; font-size: 12px; margin-top: 6px; line-height: 1.4; white-space: pre-wrap;">${Utils.escapeHtml(e.body)}</div>
                                            <button onclick="var s=document.getElementById('email_short_${cuid}_${eIdx}');var f=document.getElementById('email_full_${cuid}_${eIdx}');if(f.style.display==='none'){f.style.display='block';s.style.display='none';this.textContent='▲ Réduire';}else{f.style.display='none';s.style.display='block';this.textContent='▼ Voir plus';}" style="background: #1976d2; color: white; border: none; padding: 4px 10px; border-radius: 4px; font-size: 11px; cursor: pointer; margin-top: 6px;">▼ Voir plus</button>
                                        ` : ''}
                                    ` : ''}
                                </div>
                            `).join('')}
                        </div>
                        ` : ''}

                        ${client.emailsAffected.length > 0 ? `
                        <div style="background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%); border-radius: 10px; padding: 15px; border-left: 4px solid #388e3c;">
                            <div style="font-weight: 600; color: #2e7d32; margin-bottom: 10px; font-size: 14px;">📥 Emails reçus/affectés (${client.emailsAffected.length})</div>
                            ${client.emailsAffected.map(e => `
                                <div style="background: white; border-radius: 6px; padding: 10px; margin-bottom: 6px;">
                                    <strong style="color: #333; font-size: 13px;">${Utils.escapeHtml(e.subject || 'Sans objet')}</strong>
                                    <div style="color: #666; font-size: 11px; margin-top: 4px;">De: ${Utils.escapeHtml(e.from || '')} → ${Utils.escapeHtml(e.affectedTo || '')}</div>
                                </div>
                            `).join('')}
                        </div>
                        ` : ''}

                        ${client.aircallCalls.length > 0 ? `
                        <div style="background: linear-gradient(135deg, #fff3e0 0%, #ffe0b2 100%); border-radius: 10px; padding: 15px; border-left: 4px solid #f57c00;">
                            <div style="font-weight: 600; color: #e65100; margin-bottom: 10px; font-size: 14px;">📞 Appels téléphoniques (${client.aircallCalls.length})</div>
                            ${client.aircallCalls.map(call => {
                                const bgColor = call.type === 'sortant' ? '#fff8e1' : '#e8f5e9';
                                const borderColor = call.type === 'sortant' ? '#ffb300' : '#66bb6a';
                                const moodIcon = call.mood === 'Positif' ? '😊' : (call.mood === 'Négatif' ? '😟' : (call.mood === 'Neutre' ? '😐' : ''));
                                return `
                                <div style="background: ${bgColor}; border-radius: 8px; padding: 12px; margin-bottom: 8px; border-left: 3px solid ${borderColor};">
                                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                                        <span style="font-weight: 600; color: #333;">
                                            ${call.type === 'sortant' ? '📤 Sortant' : '📥 Entrant'}
                                            <span style="font-weight: normal; color: #666;">• ${call.duration || ''}</span>
                                            ${moodIcon ? `<span style="margin-left: 8px;">${moodIcon}</span>` : ''}
                                        </span>
                                        <span style="color: #888; font-size: 11px;">${call.time || ''}</span>
                                    </div>
                                    ${call.summary ? `
                                    <div style="background: white; border-radius: 6px; padding: 10px; font-size: 12px; color: #555; line-height: 1.5;">
                                        <div style="color: #ff8f00; font-size: 10px; font-weight: 600; margin-bottom: 4px;">💬 RÉSUMÉ IA</div>
                                        ${Utils.escapeHtml(call.summary)}
                                    </div>
                                    ` : ''}
                                </div>
                                `;
                            }).join('')}
                        </div>
                        ` : ''}

                        ${client.tasksCompleted.length > 0 ? `
                        <div style="background: linear-gradient(135deg, #fff8e1 0%, #ffecb3 100%); border-radius: 10px; padding: 15px; border-left: 4px solid #ff8f00;">
                            <div style="font-weight: 600; color: #e65100; margin-bottom: 10px; font-size: 14px;">✅ Tâches terminées (${client.tasksCompleted.length})</div>
                            ${client.tasksCompleted.map((t, tIdx) => `
                                <div style="background: white; border-radius: 6px; padding: 10px; margin-bottom: 6px;">
                                    <div style="display: flex; justify-content: space-between; align-items: flex-start;">
                                        <strong style="color: #333; font-size: 13px;">${Utils.escapeHtml(t.title)}</strong>
                                        ${t.closedTime ? `<span style="background: #ff8f00; color: white; padding: 2px 8px; border-radius: 10px; font-size: 10px;">⏰ ${t.closedTime}</span>` : ''}
                                    </div>
                                    ${t.content ? `
                                        <div id="task_short_${cuid}_${tIdx}" style="color: #666; font-size: 12px; margin-top: 6px; line-height: 1.4;">${Utils.escapeHtml(Utils.truncate(t.content, 120))}</div>
                                        ${t.content.length > 120 ? `
                                            <div id="task_full_${cuid}_${tIdx}" style="display: none; color: #666; font-size: 12px; margin-top: 6px; line-height: 1.4; white-space: pre-wrap;">${Utils.escapeHtml(t.content)}</div>
                                            <button onclick="var s=document.getElementById('task_short_${cuid}_${tIdx}');var f=document.getElementById('task_full_${cuid}_${tIdx}');if(f.style.display==='none'){f.style.display='block';s.style.display='none';this.textContent='▲ Réduire';}else{f.style.display='none';s.style.display='block';this.textContent='▼ Voir plus';}" style="background: #ff8f00; color: white; border: none; padding: 4px 10px; border-radius: 4px; font-size: 11px; cursor: pointer; margin-top: 6px;">▼ Voir plus</button>
                                        ` : ''}
                                    ` : ''}
                                </div>
                            `).join('')}
                        </div>
                        ` : ''}

                        ${client.tasksOverdue.length > 0 ? `
                        <div style="background: linear-gradient(135deg, #ffebee 0%, #ffcdd2 100%); border-radius: 10px; padding: 15px; border-left: 4px solid #d32f2f;">
                            <div style="font-weight: 600; color: #c62828; margin-bottom: 10px; font-size: 14px;">⚠️ Tâches en retard (${client.tasksOverdue.length})</div>
                            ${client.tasksOverdue.map(t => `
                                <div style="background: white; border-radius: 6px; padding: 10px; margin-bottom: 6px;">
                                    <strong style="color: #333; font-size: 13px;">${Utils.escapeHtml(t.title)}</strong>
                                    <div style="color: #d32f2f; font-size: 11px; margin-top: 4px;">${t.daysOverdue}j de retard • → ${Utils.escapeHtml(t.assignedTo || 'N/A')}</div>
                                </div>
                            `).join('')}
                        </div>
                        ` : ''}

                        ${(client.estimates.length > 0 || client.policies.length > 0 || client.claims.length > 0) ? `
                        <div style="background: linear-gradient(135deg, #f3e5f5 0%, #e1bee7 100%); border-radius: 10px; padding: 15px; border-left: 4px solid #7b1fa2;">
                            <div style="font-weight: 600; color: #6a1b9a; margin-bottom: 10px; font-size: 14px;">📄 Documents (${client.estimates.length + client.policies.length + client.claims.length})</div>
                            <div style="display: flex; flex-wrap: wrap; gap: 8px;">
                                ${client.estimates.map(e => `<span style="background: white; color: #7b1fa2; padding: 6px 12px; border-radius: 6px; font-size: 12px;">📋 Devis ${e.entityId || ''}</span>`).join('')}
                                ${client.policies.map(p => `<span style="background: white; color: #00796b; padding: 6px 12px; border-radius: 6px; font-size: 12px;">📄 Contrat ${p.entityId || ''}</span>`).join('')}
                                ${client.claims.map(c => `<span style="background: white; color: #c62828; padding: 6px 12px; border-radius: 6px; font-size: 12px;">🚨 Sinistre ${c.entityId || ''}</span>`).join('')}
                            </div>
                        </div>
                        ` : ''}
                    </div>
                </div>
                `;
            });

            html += '</div>';
            return html;
        },

        // Générer les détails d'un client pour l'export HTML
        generateClientDetailsHTML(client, clientIdx) {
            let html = '<div style="display: grid; gap: 12px;">';
            const cuid = 'exp_' + clientIdx;

            if (client.emailsSent.length > 0) {
                html += `<div style="background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%); border-radius: 10px; padding: 15px; border-left: 4px solid #1976d2;">
                    <div style="font-weight: 600; color: #1565c0; margin-bottom: 10px; font-size: 14px;">📤 Emails envoyés (${client.emailsSent.length})</div>
                    ${client.emailsSent.map((e, eIdx) => `
                        <div style="background: white; border-radius: 6px; padding: 10px; margin-bottom: 6px;">
                            <div style="display: flex; justify-content: space-between;">
                                <strong style="color: #333; font-size: 13px;">${Utils.escapeHtml(e.subject || 'Sans objet')}</strong>
                                <span style="color: #1976d2; font-size: 11px; font-weight: 500;">${e.time || ''}</span>
                            </div>
                            ${e.body ? `
                                <div id="email_short_${cuid}_${eIdx}" style="color: #666; font-size: 12px; margin-top: 6px; line-height: 1.4;">${Utils.escapeHtml(Utils.truncate(e.body, 150))}</div>
                                ${e.body.length > 150 ? `
                                    <div id="email_full_${cuid}_${eIdx}" style="display: none; color: #666; font-size: 12px; margin-top: 6px; line-height: 1.4; white-space: pre-wrap;">${Utils.escapeHtml(e.body)}</div>
                                    <button onclick="var s=document.getElementById('email_short_${cuid}_${eIdx}');var f=document.getElementById('email_full_${cuid}_${eIdx}');if(f.style.display==='none'){f.style.display='block';s.style.display='none';this.textContent='▲ Réduire';}else{f.style.display='none';s.style.display='block';this.textContent='▼ Voir plus';}" style="background: #1976d2; color: white; border: none; padding: 4px 10px; border-radius: 4px; font-size: 11px; cursor: pointer; margin-top: 6px;">▼ Voir plus</button>
                                ` : ''}
                            ` : ''}
                        </div>
                    `).join('')}
                </div>`;
            }

            if (client.emailsAffected.length > 0) {
                html += `<div style="background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%); border-radius: 10px; padding: 15px; border-left: 4px solid #388e3c;">
                    <div style="font-weight: 600; color: #2e7d32; margin-bottom: 10px; font-size: 14px;">📥 Emails reçus/affectés (${client.emailsAffected.length})</div>
                    ${client.emailsAffected.map(e => `
                        <div style="background: white; border-radius: 6px; padding: 10px; margin-bottom: 6px;">
                            <strong style="color: #333; font-size: 13px;">${Utils.escapeHtml(e.subject || 'Sans objet')}</strong>
                            <div style="color: #666; font-size: 11px; margin-top: 4px;">De: ${Utils.escapeHtml(e.from || '')} → ${Utils.escapeHtml(e.affectedTo || '')}</div>
                        </div>
                    `).join('')}
                </div>`;
            }

            if (client.aircallCalls && client.aircallCalls.length > 0) {
                html += `<div style="background: linear-gradient(135deg, #fff3e0 0%, #ffe0b2 100%); border-radius: 10px; padding: 15px; border-left: 4px solid #f57c00;">
                    <div style="font-weight: 600; color: #e65100; margin-bottom: 10px; font-size: 14px;">📞 Appels téléphoniques (${client.aircallCalls.length})</div>
                    ${client.aircallCalls.map(call => {
                        const bgColor = call.type === 'sortant' ? '#fff8e1' : '#e8f5e9';
                        const borderColor = call.type === 'sortant' ? '#ffb300' : '#66bb6a';
                        const moodIcon = call.mood === 'Positif' ? '😊' : (call.mood === 'Négatif' ? '😟' : (call.mood === 'Neutre' ? '😐' : ''));
                        return `
                        <div style="background: ${bgColor}; border-radius: 8px; padding: 12px; margin-bottom: 8px; border-left: 3px solid ${borderColor};">
                            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                                <span style="font-weight: 600; color: #333;">
                                    ${call.type === 'sortant' ? '📤 Sortant' : '📥 Entrant'}
                                    <span style="font-weight: normal; color: #666;">• ${call.duration || ''}</span>
                                    ${moodIcon ? `<span style="margin-left: 8px;">${moodIcon}</span>` : ''}
                                </span>
                                <span style="color: #888; font-size: 11px;">${call.time || ''}</span>
                            </div>
                            ${call.summary ? `
                            <div style="background: white; border-radius: 6px; padding: 10px; font-size: 12px; color: #555; line-height: 1.5;">
                                <div style="color: #ff8f00; font-size: 10px; font-weight: 600; margin-bottom: 4px;">💬 RÉSUMÉ IA</div>
                                ${Utils.escapeHtml(call.summary)}
                            </div>
                            ` : ''}
                        </div>
                        `;
                    }).join('')}
                </div>`;
            }

            if (client.tasksCompleted.length > 0) {
                html += `<div style="background: linear-gradient(135deg, #fff8e1 0%, #ffecb3 100%); border-radius: 10px; padding: 15px; border-left: 4px solid #ff8f00;">
                    <div style="font-weight: 600; color: #e65100; margin-bottom: 10px; font-size: 14px;">✅ Tâches terminées (${client.tasksCompleted.length})</div>
                    ${client.tasksCompleted.map((t, tIdx) => `
                        <div style="background: white; border-radius: 6px; padding: 10px; margin-bottom: 6px;">
                            <div style="display: flex; justify-content: space-between; align-items: flex-start;">
                                <strong style="color: #333; font-size: 13px;">${Utils.escapeHtml(t.title)}</strong>
                                ${t.closedTime ? `<span style="background: #ff8f00; color: white; padding: 2px 8px; border-radius: 10px; font-size: 10px;">⏰ ${t.closedTime}</span>` : ''}
                            </div>
                            ${t.content ? `
                                <div id="task_short_${cuid}_${tIdx}" style="color: #666; font-size: 12px; margin-top: 6px; line-height: 1.4;">${Utils.escapeHtml(Utils.truncate(t.content, 120))}</div>
                                ${t.content.length > 120 ? `
                                    <div id="task_full_${cuid}_${tIdx}" style="display: none; color: #666; font-size: 12px; margin-top: 6px; line-height: 1.4; white-space: pre-wrap;">${Utils.escapeHtml(t.content)}</div>
                                    <button onclick="var s=document.getElementById('task_short_${cuid}_${tIdx}');var f=document.getElementById('task_full_${cuid}_${tIdx}');if(f.style.display==='none'){f.style.display='block';s.style.display='none';this.textContent='▲ Réduire';}else{f.style.display='none';s.style.display='block';this.textContent='▼ Voir plus';}" style="background: #ff8f00; color: white; border: none; padding: 4px 10px; border-radius: 4px; font-size: 11px; cursor: pointer; margin-top: 6px;">▼ Voir plus</button>
                                ` : ''}
                            ` : ''}
                        </div>
                    `).join('')}
                </div>`;
            }

            if (client.tasksOverdue.length > 0) {
                html += `<div style="background: linear-gradient(135deg, #ffebee 0%, #ffcdd2 100%); border-radius: 10px; padding: 15px; border-left: 4px solid #d32f2f;">
                    <div style="font-weight: 600; color: #c62828; margin-bottom: 10px; font-size: 14px;">⚠️ Tâches en retard (${client.tasksOverdue.length})</div>
                    ${client.tasksOverdue.map(t => `
                        <div style="background: white; border-radius: 6px; padding: 10px; margin-bottom: 6px;">
                            <strong style="color: #333; font-size: 13px;">${Utils.escapeHtml(t.title)}</strong>
                            <div style="color: #d32f2f; font-size: 11px; margin-top: 4px;">${t.daysOverdue}j de retard • → ${Utils.escapeHtml(t.assignedTo || 'N/A')}</div>
                        </div>
                    `).join('')}
                </div>`;
            }

            if (client.estimates.length > 0 || client.policies.length > 0 || client.claims.length > 0) {
                html += `<div style="background: linear-gradient(135deg, #f3e5f5 0%, #e1bee7 100%); border-radius: 10px; padding: 15px; border-left: 4px solid #7b1fa2;">
                    <div style="font-weight: 600; color: #6a1b9a; margin-bottom: 10px; font-size: 14px;">📄 Documents (${client.estimates.length + client.policies.length + client.claims.length})</div>
                    <div style="display: flex; flex-wrap: wrap; gap: 8px;">
                        ${client.estimates.map(e => `<span style="background: white; color: #7b1fa2; padding: 6px 12px; border-radius: 6px; font-size: 12px;">📋 Devis ${e.entityId || ''}</span>`).join('')}
                        ${client.policies.map(p => `<span style="background: white; color: #00796b; padding: 6px 12px; border-radius: 6px; font-size: 12px;">📄 Contrat ${p.entityId || ''}</span>`).join('')}
                        ${client.claims.map(c => `<span style="background: white; color: #c62828; padding: 6px 12px; border-radius: 6px; font-size: 12px;">🚨 Sinistre ${c.entityId || ''}</span>`).join('')}
                    </div>
                </div>`;
            }

            html += '</div>';
            return html || '<p style="color: #999;">Aucune action</p>';
        }
    };

    // ============================================
    // LOADER UI
    // ============================================
    function showLoader() {
        const reportDate = Utils.getTodayDate();
        const realToday = Utils.getRealTodayDate();
        const isPastDate = reportDate !== realToday;
        const dateDisplay = isPastDate
            ? `📅 <span style="color: #ff9800; font-weight: bold;">${reportDate}</span> <span style="background: #ff9800; color: white; padding: 2px 8px; border-radius: 3px; font-size: 11px;">Rétrospectif</span>`
            : `📅 ${reportDate}`;

        const loader = document.createElement('div');
        loader.id = 'ltoa-loader';
        loader.innerHTML = `
            <div style="
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background: rgba(0,0,0,0.9);
                z-index: 999999;
                display: flex;
                flex-direction: column;
                justify-content: center;
                align-items: center;
                color: white;
                font-family: Arial, sans-serif;
            ">
                <div style="
                    background: white;
                    border-radius: 15px;
                    padding: 40px 60px;
                    text-align: center;
                    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
                    min-width: 400px;
                ">
                    <div style="font-size: 60px; margin-bottom: 20px;" id="loader-emoji">⏳</div>
                    <h2 style="color: #333; margin: 0 0 10px 0;" id="loader-title">Génération du rapport...</h2>
                    <p style="color: #666; margin: 0 0 5px 0; font-size: 14px;">${dateDisplay}</p>
                    <p style="color: #666; margin: 0 0 20px 0; min-height: 20px;" id="loader-status">Initialisation...</p>

                    <div style="width: 300px; height: 8px; background: #e0e0e0; border-radius: 4px; overflow: hidden; margin-bottom: 30px;">
                        <div id="loader-progress" style="width: 0%; height: 100%; background: linear-gradient(90deg, #c62828, #ff5722); transition: width 0.5s ease;"></div>
                    </div>

                    <div id="loader-steps" style="text-align: left; font-size: 14px;">
                        <p style="margin: 8px 0; color: #666;" id="step-1">⏳ Emails envoyés...</p>
                        <p style="margin: 8px 0; color: #bbb;" id="step-2">⏳ Emails affectés...</p>
                        <p style="margin: 8px 0; color: #bbb;" id="step-3">⏳ Tâches terminées...</p>
                        <p style="margin: 8px 0; color: #bbb;" id="step-4">⏳ Tâches en retard...</p>
                        <p style="margin: 8px 0; color: #bbb;" id="step-5">⏳ Devis...</p>
                        <p style="margin: 8px 0; color: #bbb;" id="step-6">⏳ Contrats...</p>
                        <p style="margin: 8px 0; color: #bbb;" id="step-7">⏳ Sinistres...</p>
                        <p style="margin: 8px 0; color: #bbb;" id="step-8">⏳ Autres actions...</p>
                    </div>
                </div>
            </div>
        `;
        document.body.appendChild(loader);

        return {
            update: (step, progress, status) => {
                const statusEl = document.getElementById('loader-status');
                const progressEl = document.getElementById('loader-progress');
                if (statusEl) statusEl.textContent = status;
                if (progressEl) progressEl.style.width = `${progress}%`;

                for (let i = 1; i <= 8; i++) {
                    const stepEl = document.getElementById(`step-${i}`);
                    if (stepEl) {
                        if (i < step) {
                            stepEl.innerHTML = stepEl.innerHTML.replace('⏳', '✅');
                            stepEl.style.color = '#4caf50';
                        } else if (i === step) {
                            stepEl.style.color = '#333';
                            stepEl.style.fontWeight = 'bold';
                        }
                    }
                }
            },
            updateStatus: (status) => {
                const el = document.getElementById('loader-status');
                if (el) el.textContent = status;
            },
            complete: () => {
                const emoji = document.getElementById('loader-emoji');
                const title = document.getElementById('loader-title');
                const progress = document.getElementById('loader-progress');

                if (emoji) emoji.textContent = '✅';
                if (title) title.textContent = 'Rapport prêt !';
                if (progress) progress.style.width = '100%';

                for (let i = 1; i <= 8; i++) {
                    const stepEl = document.getElementById(`step-${i}`);
                    if (stepEl) {
                        stepEl.innerHTML = stepEl.innerHTML.replace('⏳', '✅');
                        stepEl.style.color = '#4caf50';
                    }
                }
            },
            remove: () => {
                const el = document.getElementById('ltoa-loader');
                if (el) el.remove();
            }
        };
    }

    // ============================================
    // GÉNÉRATION DU RAPPORT
    // ============================================
    async function generateReport() {
        const loader = showLoader();

        try {
            const connectedUser = Utils.getConnectedUser();
            const userId = Utils.getUserData(connectedUser);
            Utils.log('Utilisateur détecté:', connectedUser, userId);

            // Étape 1: Emails envoyés
            loader.update(1, 5, 'Collecte des emails envoyés...');
            const emailsSent = await EmailsSentCollector.collect(connectedUser, loader.updateStatus);
            await Utils.delay(CONFIG.DELAY_BETWEEN_REQUESTS);

            // Étape 2: Emails affectés
            loader.update(2, 10, 'Collecte des emails affectés...');
            const emailsAffected = await EmailsAffectedCollector.collect(connectedUser, loader.updateStatus);
            await Utils.delay(CONFIG.DELAY_BETWEEN_REQUESTS);

            // Étape 2b: Nombre d'emails en attente
            loader.update(2, 15, 'Comptage emails en attente...');
            const pendingEmailsCount = await PendingEmailsCollector.collect(connectedUser, loader.updateStatus);
            await Utils.delay(CONFIG.DELAY_BETWEEN_REQUESTS);

            // Étape 3: Appels Aircall
            loader.update(3, 20, 'Collecte des appels Aircall...');
            let aircallCalls = [];
            if (CONFIG.AIRCALL_ENABLED) {
                try {
                    aircallCalls = await AircallCollector.collect(connectedUser, loader.updateStatus);
                    Utils.log(`${aircallCalls.length} appels Aircall collectés`);
                } catch (e) {
                    Utils.log('Erreur collecte Aircall (non bloquante):', e);
                }
            }
            await Utils.delay(CONFIG.DELAY_BETWEEN_REQUESTS);

            // Étape 4: Tâches terminées
            loader.update(4, 35, 'Collecte des tâches terminées...');
            const tasksCompleted = await TasksCompletedCollector.collect(userId, connectedUser, loader.updateStatus);
            await Utils.delay(CONFIG.DELAY_BETWEEN_REQUESTS);

            // Étape 5: Tâches en retard
            loader.update(5, 48, 'Collecte des tâches en retard...');
            const tasksOverdue = await TasksOverdueCollector.collect(userId, connectedUser, loader.updateStatus);
            await Utils.delay(CONFIG.DELAY_BETWEEN_REQUESTS);

            // Étape 6: Devis
            loader.update(6, 58, 'Collecte des devis...');
            const estimates = await LogsCollector.collectEstimates(userId, connectedUser, loader.updateStatus);
            await Utils.delay(CONFIG.DELAY_BETWEEN_REQUESTS);

            // Étape 7: Contrats
            loader.update(7, 72, 'Collecte des contrats...');
            const policies = await LogsCollector.collectPolicies(userId, connectedUser, loader.updateStatus);
            await Utils.delay(CONFIG.DELAY_BETWEEN_REQUESTS);

            // Étape 8: Sinistres
            loader.update(8, 82, 'Collecte des sinistres...');
            const claims = await LogsCollector.collectClaims(userId, connectedUser, loader.updateStatus);
            await Utils.delay(CONFIG.DELAY_BETWEEN_REQUESTS);

            // Étape 9: Autres actions (journalisation générale)
            loader.update(9, 88, 'Collecte des autres actions...');
            const logs = await LogsCollector.collect(userId, connectedUser, loader.updateStatus);

            // Étape 10: Résolution des clients (correspondance email <-> N° client <-> nom)
            loader.update(10, 94, 'Résolution des clients...');
            ClientResolver.reset(); // Réinitialiser pour un nouveau rapport
            const resolvedData = await ClientResolver.resolve({
                emailsSent,
                emailsAffected,
                tasksCompleted,
                tasksOverdue,
                logs,
                estimates,
                policies,
                claims
            }, loader.updateStatus);

            // Finalisation
            loader.complete();
            await Utils.delay(800);

            ReportGenerator.data = {
                emailsSent: resolvedData.emailsSent,
                emailsAffected: resolvedData.emailsAffected,
                pendingEmailsCount: pendingEmailsCount, // Nombre d'emails en attente
                aircallCalls: aircallCalls, // Ajouter les appels Aircall
                tasksCompleted: resolvedData.tasksCompleted,
                tasksOverdue: resolvedData.tasksOverdue,
                logs: resolvedData.logs,
                estimates: resolvedData.estimates,
                policies: resolvedData.policies,
                claims: resolvedData.claims,
                notes: REPORT_NOTES,
                aircallStatus: AircallCollector.lastStatus,
                user: connectedUser,
                date: Utils.getTodayDate(),
                clientIndex: ClientResolver.clientIndex // Garder l'index pour la vue par client
            };

            loader.remove();
            ReportGenerator.show();

        } catch (error) {
            Utils.log('Erreur génération rapport:', error);
            loader.remove();
            alert(`❌ Erreur lors de la génération:\n${error.message}\n\nVoir console pour détails.`);
        }
    }

    // ============================================
    // BOUTON PRINCIPAL
    // ============================================
    let reportGenerated = false;

    // Afficher le sélecteur de date
    function showDatePicker() {
        return new Promise((resolve) => {
            // Supprimer un éventuel picker existant
            const existing = document.getElementById('ltoa-date-picker-modal');
            if (existing) existing.remove();

            // Calculer les dates pour les boutons rapides
            const today = new Date();
            const formatDate = (d) => {
                const day = String(d.getDate()).padStart(2, '0');
                const month = String(d.getMonth() + 1).padStart(2, '0');
                const year = d.getFullYear();
                return `${day}/${month}/${year}`;
            };
            const formatDateInput = (d) => {
                const day = String(d.getDate()).padStart(2, '0');
                const month = String(d.getMonth() + 1).padStart(2, '0');
                const year = d.getFullYear();
                return `${year}-${month}-${day}`;
            };

            const yesterday = new Date(today);
            yesterday.setDate(yesterday.getDate() - 1);

            const twoDaysAgo = new Date(today);
            twoDaysAgo.setDate(twoDaysAgo.getDate() - 2);

            const modal = document.createElement('div');
            modal.id = 'ltoa-date-picker-modal';
            modal.innerHTML = `
                <div style="
                    position: fixed;
                    top: 0; left: 0; right: 0; bottom: 0;
                    background: rgba(0,0,0,0.5);
                    z-index: 2147483646;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                ">
                    <div style="
                        background: white;
                        border-radius: 12px;
                        padding: 25px;
                        min-width: 350px;
                        box-shadow: 0 10px 40px rgba(0,0,0,0.3);
                        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                    ">
                        <h3 style="margin: 0 0 20px 0; color: #333; font-size: 18px; display: flex; align-items: center; gap: 10px;">
                            📅 Choisir la date du rapport
                        </h3>

                        <div style="margin-bottom: 20px;">
                            <label style="display: block; margin-bottom: 8px; color: #666; font-size: 13px;">
                                Sélectionner une date :
                            </label>
                            <input type="date" id="ltoa-date-input" value="${formatDateInput(today)}" max="${formatDateInput(today)}" style="
                                width: 100%;
                                padding: 12px;
                                border: 2px solid #e0e0e0;
                                border-radius: 8px;
                                font-size: 15px;
                                box-sizing: border-box;
                                transition: border-color 0.2s;
                            ">
                        </div>

                        <div style="margin-bottom: 20px;">
                            <label style="display: block; margin-bottom: 8px; color: #666; font-size: 13px;">
                                Raccourcis :
                            </label>
                            <div style="display: flex; gap: 8px; flex-wrap: wrap;">
                                <button class="ltoa-quick-date" data-date="${formatDate(today)}" style="
                                    flex: 1;
                                    padding: 10px;
                                    border: 2px solid #c62828;
                                    background: #fff;
                                    border-radius: 6px;
                                    cursor: pointer;
                                    font-size: 13px;
                                    color: #c62828;
                                    font-weight: 500;
                                    transition: all 0.2s;
                                ">
                                    📆 Aujourd'hui<br>
                                    <small style="color: #888;">${formatDate(today)}</small>
                                </button>
                                <button class="ltoa-quick-date" data-date="${formatDate(yesterday)}" style="
                                    flex: 1;
                                    padding: 10px;
                                    border: 2px solid #1976d2;
                                    background: #fff;
                                    border-radius: 6px;
                                    cursor: pointer;
                                    font-size: 13px;
                                    color: #1976d2;
                                    font-weight: 500;
                                    transition: all 0.2s;
                                ">
                                    ⏪ Hier<br>
                                    <small style="color: #888;">${formatDate(yesterday)}</small>
                                </button>
                                <button class="ltoa-quick-date" data-date="${formatDate(twoDaysAgo)}" style="
                                    flex: 1;
                                    padding: 10px;
                                    border: 2px solid #7b1fa2;
                                    background: #fff;
                                    border-radius: 6px;
                                    cursor: pointer;
                                    font-size: 13px;
                                    color: #7b1fa2;
                                    font-weight: 500;
                                    transition: all 0.2s;
                                ">
                                    ⏪⏪ Avant-hier<br>
                                    <small style="color: #888;">${formatDate(twoDaysAgo)}</small>
                                </button>
                            </div>
                        </div>

                        <div style="margin-bottom: 20px;">
                            <label for="ltoa-report-notes" style="display:block;margin-bottom:8px;color:#666;font-size:13px;">
                                Notes ou précisions complémentaires (facultatif) :
                            </label>
                            <textarea id="ltoa-report-notes" rows="4" placeholder="Ex. rendez-vous extérieur, travail de fond, incident technique, précision sur un dossier…" style="
                                width:100%;padding:12px;border:2px solid #e0e0e0;border-radius:8px;
                                font-size:14px;box-sizing:border-box;resize:vertical;font-family:inherit;
                            ">${Utils.escapeHtml(REPORT_NOTES)}</textarea>
                        </div>

                        <div style="margin-bottom:20px;padding:12px 14px;background:#f5f7fa;border:1px solid #e0e0e0;border-radius:8px;">
                            <div style="font-size:13px;font-weight:600;color:#333;">Connexion Aircall</div>
                            <div style="font-size:11px;color:#777;margin-top:3px;">Sécurisée via Netlify · aucune configuration nécessaire</div>
                        </div>

                        <div style="display: flex; gap: 10px; margin-top: 25px;">
                            <button id="ltoa-date-cancel" style="
                                flex: 1;
                                padding: 12px;
                                border: 2px solid #ccc;
                                background: #fff;
                                border-radius: 8px;
                                cursor: pointer;
                                font-size: 14px;
                                color: #666;
                                transition: all 0.2s;
                            ">Annuler</button>
                            <button id="ltoa-date-confirm" style="
                                flex: 1;
                                padding: 12px;
                                border: none;
                                background: linear-gradient(135deg, #c62828, #b71c1c);
                                border-radius: 8px;
                                cursor: pointer;
                                font-size: 14px;
                                color: white;
                                font-weight: 500;
                                transition: all 0.2s;
                            ">📊 Générer le rapport</button>
                        </div>
                    </div>
                </div>
            `;

            document.body.appendChild(modal);

            // Focus sur l'input date
            const dateInput = document.getElementById('ltoa-date-input');
            dateInput.focus();

            // Styles hover pour les boutons
            const styleHover = document.createElement('style');
            styleHover.textContent = `
                .ltoa-quick-date:hover {
                    transform: translateY(-2px);
                    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
                }
                #ltoa-date-input:focus {
                    border-color: #c62828;
                    outline: none;
                }
                #ltoa-date-cancel:hover {
                    background: #f5f5f5;
                }
                #ltoa-date-confirm:hover {
                    transform: translateY(-1px);
                    box-shadow: 0 4px 12px rgba(198, 40, 40, 0.4);
                }
            `;
            modal.appendChild(styleHover);

            // Événements boutons rapides
            modal.querySelectorAll('.ltoa-quick-date').forEach(btn => {
                btn.addEventListener('click', () => {
                    const date = btn.getAttribute('data-date');
                    // Convertir en format input (YYYY-MM-DD)
                    const parts = date.split('/');
                    dateInput.value = `${parts[2]}-${parts[1]}-${parts[0]}`;
                    // Effet visuel
                    modal.querySelectorAll('.ltoa-quick-date').forEach(b => {
                        b.style.background = '#fff';
                        b.style.fontWeight = '500';
                    });
                    btn.style.background = btn.style.borderColor;
                    btn.style.color = '#fff';
                });
            });

            // Annuler
            document.getElementById('ltoa-date-cancel').addEventListener('click', () => {
                modal.remove();
                resolve(null);
            });

            // Confirmer
            document.getElementById('ltoa-date-confirm').addEventListener('click', () => {
                const inputValue = dateInput.value; // Format YYYY-MM-DD
                if (inputValue) {
                    REPORT_NOTES = document.getElementById('ltoa-report-notes')?.value.trim() || '';
                    const parts = inputValue.split('-');
                    const formattedDate = `${parts[2]}/${parts[1]}/${parts[0]}`; // DD/MM/YYYY
                    modal.remove();
                    resolve(formattedDate);
                }
            });

            // Fermer en cliquant en dehors
            modal.querySelector(':first-child').addEventListener('click', (e) => {
                if (e.target === e.currentTarget) {
                    modal.remove();
                    resolve(null);
                }
            });

            // Touche Entrée pour confirmer, Echap pour annuler
            modal.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    document.getElementById('ltoa-date-confirm').click();
                } else if (e.key === 'Escape') {
                    modal.remove();
                    resolve(null);
                }
            });
        });
    }

    async function handleReportClick() {
        const existingModal = document.getElementById('ltoa-report-modal');

        if (existingModal) {
            if (existingModal.style.display === 'none') {
                existingModal.style.display = 'block';
            } else {
                const action = confirm('📊 Le rapport est déjà ouvert.\n\nOK = Générer un NOUVEAU rapport\nAnnuler = Fermer le rapport actuel');
                if (action) {
                    existingModal.remove();
                    reportGenerated = false;
                    // Afficher le sélecteur de date
                    const selectedDate = await showDatePicker();
                    if (selectedDate) {
                        SELECTED_REPORT_DATE = selectedDate;
                        Utils.log('Date sélectionnée pour le rapport:', SELECTED_REPORT_DATE);
                        generateReport();
                    }
                } else {
                    existingModal.style.display = 'none';
                }
            }
        } else {
            // Afficher le sélecteur de date
            const selectedDate = await showDatePicker();
            if (selectedDate) {
                SELECTED_REPORT_DATE = selectedDate;
                Utils.log('Date sélectionnée pour le rapport:', SELECTED_REPORT_DATE);
                generateReport();
            }
        }
    }

    function addReportButton() {
        if (!window.location.href.includes('courtage.modulr.fr')) return;
        if (document.getElementById('ltoa-daily-report-v4-btn')) return;

        // Créer le bouton dans le même style que les icônes Modulr
        const button = document.createElement('a');
        button.id = 'ltoa-daily-report-v4-btn';
        button.href = '#';
        button.className = 'left banner_icon';
        button.title = 'Rapport du Jour';
        button.style.cssText = 'cursor: pointer; text-decoration: none;';
        button.innerHTML = '<span class="fa fa-chart-bar"></span>';

        // Créer le badge (optionnel, on peut mettre un indicateur)
        const badge = document.createElement('a');
        badge.href = '#';
        badge.className = 'banner_badge';
        badge.title = 'Générer le rapport';
        badge.style.cssText = 'cursor: pointer; background: #c62828 !important;';
        badge.textContent = '📊';

        // Chercher la zone left dans le header nav
        const headerNavLeft = document.querySelector('#main-header-nav .content .left');

        if (headerNavLeft) {
            headerNavLeft.appendChild(button);
            headerNavLeft.appendChild(badge);
            Utils.log('Bouton ajouté dans header nav left (style Modulr)');
        } else {
            // Fallback: position fixe
            const fallbackBtn = document.createElement('div');
            fallbackBtn.id = 'ltoa-daily-report-v4-btn';
            fallbackBtn.innerHTML = `
                <button style="
                    position: fixed;
                    top: 8px;
                    left: 350px;
                    z-index: 2147483647;
                    background: #c62828;
                    color: white;
                    border: none;
                    padding: 5px 10px;
                    border-radius: 3px;
                    cursor: pointer;
                    font-size: 12px;
                ">📊 Rapport</button>
            `;
            document.body.appendChild(fallbackBtn);
            fallbackBtn.querySelector('button').addEventListener('click', handleReportClick);
            Utils.log('Bouton ajouté en position fixe (fallback)');
            return;
        }

        // Event listeners
        button.addEventListener('click', (e) => {
            e.preventDefault();
            handleReportClick();
        });
        badge.addEventListener('click', (e) => {
            e.preventDefault();
            handleReportClick();
        });

        Utils.log('Bouton rapport V4 ajouté avec succès (style Modulr) !');
    }

    // ============================================
    // INITIALISATION
    // ============================================
    function init() {
        Utils.log('Script LTOA Rapport V4 chargé');

        // Attendre que le DOM soit prêt
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', () => {
                setTimeout(addReportButton, 1000);
            });
        } else {
            setTimeout(addReportButton, 1000);
        }

        // Observer pour ré-ajouter le bouton si supprimé
        const observer = new MutationObserver(() => {
            if (!document.getElementById('ltoa-daily-report-v4-btn')) {
                addReportButton();
            }
        });

        setTimeout(() => {
            observer.observe(document.body, { childList: true, subtree: true });
        }, 2000);
    }

    init();

})();
