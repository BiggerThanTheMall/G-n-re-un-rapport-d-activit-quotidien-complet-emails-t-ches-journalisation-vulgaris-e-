// ==UserScript==
// @name         LTOA Modulr - Rapport Quotidien
// @namespace    https://github.com/BiggerThanTheMall/tampermonkey-ltoa
// @version      5.6.4
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
            'delivered': 'Remis',
            'pending_parts': 'Attente pièces',
            'pending_approval': 'Attente approbation',
            'deferred': 'Différé',
            'lost': 'Perdu',
            'transformed': 'Transformé',
            'subscription': 'En souscription',
            'underwriting': 'En souscription',
            'in_subscription': 'En souscription',
            'accepted': 'Accepté',
            'refused': 'Refusé',
            'expired': 'Expiré',
            'canceled': 'Annulé',
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

        cleanRichText: (text) => {
            if (!text) return '';

            let value = String(text)
                .replace(/\\r\\n/g, '\n')
                .replace(/\\n/g, '\n')
                .replace(/\\r/g, '\n')
                .replace(/\\t/g, ' ')
                .replace(/\\\//g, '/');

            // Certaines notes Modulr sont encodées plusieurs fois :
            // &amp;#x20; -> &#x20; -> espace. On décode jusqu'à stabilisation.
            const decoder = document.createElement('textarea');
            for (let i = 0; i < 5; i++) {
                const before = value;
                decoder.innerHTML = value;
                value = decoder.value;
                value = value
                    .replace(/&#x([0-9a-f]+);/gi, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
                    .replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(parseInt(dec, 10)));
                if (value === before) break;
            }

            value = value
                .replace(/<br\s*\/?>/gi, '\n')
                .replace(/<\/p>/gi, '\n')
                .replace(/<\/div>/gi, '\n')
                .replace(/<\/li>/gi, '\n')
                .replace(/<li[^>]*>/gi, '• ')
                .replace(/<[^>]+>/g, '')
                .replace(/&nbsp;/gi, ' ');

            // Dernier passage après suppression des balises.
            for (let i = 0; i < 3; i++) {
                const before = value;
                decoder.innerHTML = value;
                value = decoder.value;
                if (value === before) break;
            }

            return value
                .replace(/\\t/g, ' ')
                .replace(/\r\n/g, '\n')
                .replace(/\r/g, '\n')
                .replace(/[\u00a0\u2000-\u200b\u202f\u205f\u3000\ufeff]/g, ' ')
                .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '')
                .replace(/[ \t]+/g, ' ')
                .replace(/ *\n */g, '\n')
                .replace(/\n{3,}/g, '\n\n')
                .trim();
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
            return TRANSLATIONS.values[strValue] || TRANSLATIONS.values[strValue.toLowerCase()] || strValue;
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
        normalizeName(value) {
            return String(value || '')
                .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
                .toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
        },

        async collect(connectedUser, updateLoader) {
            Utils.log('=== COLLECTE EMAILS EN ATTENTE ===');
            Utils.log('Utilisateur recherché:', connectedUser);

            try {
                if (typeof updateLoader === 'function') updateLoader('Récupération emails en attente...');

                const url = 'https://courtage.modulr.fr/fr/scripts/emails/emails_list.php?email_page=1';
                const html = await Utils.fetchPage(url);
                const doc = Utils.parseHTML(html);
                const wanted = this.normalizeName(connectedUser);
                const wantedParts = wanted.split(' ').filter(Boolean);

                // Le menu d'affectation contient le compteur de référence de Modulr.
                // On cherche d'abord un nom COMPLET afin d'éviter de confondre Eddy,
                // Nadia, Doryan ou Ghaïs simplement parce qu'ils partagent "Kalah".
                const candidates = Array.from(doc.querySelectorAll('a, button, li, span'))
                    .map(el => ({
                        el,
                        text: (el.textContent || '').replace(/\s+/g, ' ').trim()
                    }))
                    .filter(item => /\(\s*\d+\s*\)\s*$/.test(item.text));

                const parsed = [];
                for (const item of candidates) {
                    const match = item.text.match(/^(.+?)\s*\(\s*(\d+)\s*\)\s*$/);
                    if (!match) continue;
                    parsed.push({
                        name: match[1].trim(),
                        normalized: this.normalizeName(match[1]),
                        count: parseInt(match[2], 10) || 0
                    });
                }

                let found = parsed.find(item => item.normalized === wanted);

                // Fallback tolérant aux variantes d'affichage, mais il faut que TOUS
                // les éléments du nom correspondent. Jamais de match sur le seul nom de famille.
                if (!found && wantedParts.length) {
                    found = parsed.find(item => {
                        const parts = item.normalized.split(' ').filter(Boolean);
                        return wantedParts.every(part => parts.includes(part)) &&
                               parts.every(part => wantedParts.includes(part));
                    });
                }

                if (!found) {
                    Utils.log('Compteur email utilisateur introuvable. Candidats:', parsed);
                    return 0;
                }

                Utils.log(`✓ ${found.name} : ${found.count} emails actuellement affectés`);
                return found.count;
            } catch (error) {
                Utils.log('Erreur collecte emails en attente:', error);
                return 0;
            }
        }
    };

    // ============================================
    // COLLECTEUR D'APPELS AIRCALL
    // ============================================
    // Accès Aircall sécurisé via un relais serveur : aucun secret dans Tampermonkey.
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
                updateLoader(`Aircall : page ${page}...`);
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
                updateLoader('Connexion à Aircall...');
                const calls = await this.collectDirect(connectedUser, reportDate, updateLoader);
                this.lastStatus = { state: 'complete', source: 'api', message: `Aircall : ${calls.length} appel${calls.length > 1 ? 's' : ''} trouvé${calls.length > 1 ? 's' : ''}` };
                return calls;
            } catch (error) {
                this.lastStatus = { state: 'error', source: 'api', message: `Erreur API Aircall : ${error.message}` };
                throw error;
            }
        }
    };

    // ============================================
    // COLLECTEUR DES DEVIS ACTUELLEMENT ASSIGNÉS
    // ============================================
    const EstimatesAssignmentCollector = {
        async collect(userId, updateLoader) {
            const empty = { total: 0, statuses: [] };
            try {
                if (updateLoader) updateLoader('Comptage des devis actuellement assignés...');
                const html = await Utils.fetchPage('https://courtage.modulr.fr/fr/scripts/dashboard/dashboard.php');
                const doc = Utils.parseHTML(html);
                const referentId = String(userId?.logValue || '').trim();
                if (!referentId) return empty;

                const tables = Array.from(doc.querySelectorAll('table.box_statistics'));
                const table = tables.find(t => /Mes devis/i.test(t.textContent || ''));
                if (!table) {
                    Utils.log('Widget "Mes devis" introuvable sur le tableau de bord');
                    return empty;
                }

                const statuses = [];
                for (const row of table.querySelectorAll('tr.box_statistics_line')) {
                    const label = row.querySelector('.box_statistics_status')?.textContent?.replace(/\s+/g, ' ').trim();
                    if (!label || /sans référent/i.test(label)) continue;

                    const links = Array.from(row.querySelectorAll('a[href*="referent_user_id="]'));
                    const mine = links.find(a => {
                        try {
                            const url = new URL(a.getAttribute('href'), location.origin);
                            return url.searchParams.get('referent_user_id') === referentId &&
                                   url.searchParams.get('no_referent') !== '1';
                        } catch (_) {
                            return false;
                        }
                    });
                    if (!mine) continue;

                    const count = parseInt(mine.querySelector('.widget_number')?.textContent?.trim() || '0', 10) || 0;
                    let status = '';
                    try {
                        const url = new URL(mine.getAttribute('href'), location.origin);
                        status = url.searchParams.get('status') || '';
                    } catch (_) {}

                    statuses.push({ label, status, count });
                }

                return {
                    total: statuses.reduce((sum, item) => sum + item.count, 0),
                    statuses
                };
            } catch (error) {
                Utils.log('Erreur comptage devis assignés:', error);
                return empty;
            }
        }
    };

    // ============================================
    // COLLECTEUR DE TÂCHES TERMINÉES
    // ============================================
    const TaskListUtils = {
        baseUrl: 'https://courtage.modulr.fr/fr/scripts/Tasks/TasksList.php',

        taskId(row) {
            const direct = row?.getAttribute?.('data-task-id');
            if (direct && /^\d+$/.test(direct)) return direct;
            const id = row?.id || '';
            const match = id.match(/^task[:_-](\d+)$/i);
            return match ? match[1] : null;
        },

        rows(doc) {
            return Array.from(doc.querySelectorAll('tr')).filter(row => this.taskId(row));
        },

        dueDate(row) {
            const cells = Array.from(row.querySelectorAll('td.align_center, td'));
            for (const cell of cells) {
                const text = (cell.textContent || '').replace(/\s+/g, ' ').trim();
                const match = text.match(/\b(\d{2}\/\d{2}\/\d{4})(?:\s+(?:à\s*)?(\d{1,2}:\d{2}))?/);
                if (match) return match[1] + (match[2] ? ` à ${match[2]}` : '');
            }
            return '';
        },

        audit(row) {
            const hidden = row.querySelector('.hidden');
            const raw = hidden ? ((hidden.textContent || '') + ' ' + (hidden.innerHTML || '')) : '';
            const normalized = raw.replace(/<[^>]+>/g, ' ').replace(/&nbsp;/gi, ' ').replace(/\s+/g, ' ').trim();

            let lastModifiedDate = '';
            let lastModifiedTime = '';
            let lastModifiedBy = '';
            const mod = normalized.match(/Derni[èe]re modification\s*:?\s*(.*?)\s+(\d{2}\/\d{2}\/\d{4})(?:\s+(\d{1,2}:\d{2})(?::\d{2})?)?/i);
            if (mod) {
                lastModifiedBy = (mod[1] || '').trim();
                lastModifiedDate = mod[2] || '';
                lastModifiedTime = mod[3] || '';
            }

            let createdDate = '';
            let createdBy = '';
            const creation = normalized.match(/Cr[ée]ation\s*:?\s*(.*?)\s+(\d{2}\/\d{2}\/\d{4})/i);
            if (creation) {
                createdBy = (creation[1] || '').trim();
                createdDate = creation[2] || '';
            }

            return { lastModifiedDate, lastModifiedTime, lastModifiedBy, createdDate, createdBy };
        },

        isOverdue(row, dueDate) {
            const byStyle = row.classList.contains('task_late_background_color') ||
                !!row.querySelector('.task_late_icon, .task_late_divider, .fa-exclamation-triangle, [class*="late"], [class*="overdue"]');
            return byStyle || TasksOverdueCollector.calculateDaysOverdue(dueDate) > 0;
        },

        nextPageUrl(doc, currentUrl, visited) {
            const links = Array.from(doc.querySelectorAll('a[href]'));
            const direct = links.find(a => {
                const hint = [
                    a.getAttribute('rel') || '',
                    a.getAttribute('title') || '',
                    a.getAttribute('aria-label') || '',
                    a.textContent || ''
                ].join(' ').replace(/\s+/g, ' ').trim();
                return /(^|\s)(next|suivant|page suivante)(\s|$)|^[›»>]$/i.test(hint);
            });

            const normalize = href => {
                try { return new URL(href, currentUrl).href; } catch (_) { return null; }
            };

            if (direct) {
                const url = normalize(direct.getAttribute('href'));
                if (url && !visited.has(url)) return url;
            }

            // Fallback : Modulr change parfois le nom du paramètre de pagination.
            // On détecte donc tout paramètre contenant "page".
            const numeric = [];
            for (const a of links) {
                const url = normalize(a.getAttribute('href'));
                if (!url || visited.has(url)) continue;
                try {
                    const parsed = new URL(url);
                    if (!/TasksList\.php/i.test(parsed.pathname)) continue;
                    for (const [key, value] of parsed.searchParams.entries()) {
                        if (/page/i.test(key) && /^\d+$/.test(value)) {
                            numeric.push({ url, page: parseInt(value, 10) });
                        }
                    }
                } catch (_) {}
            }
            numeric.sort((a, b) => a.page - b.page);
            return numeric.length ? numeric[0].url : null;
        },

        async fetchAll(params, updateLoader, label) {
            let currentUrl = `${this.baseUrl}?${params.toString()}`;
            const visited = new Set();
            const pages = [];

            for (let page = 1; page <= 30 && currentUrl && !visited.has(currentUrl); page++) {
                visited.add(currentUrl);
                if (typeof updateLoader === 'function') updateLoader(`${label} - page ${page}...`);
                const html = await Utils.fetchPage(currentUrl);
                const doc = Utils.parseHTML(html);
                const rows = this.rows(doc);
                pages.push({ url: currentUrl, html, doc, rows, allRows: Array.from(doc.querySelectorAll('tr')) });

                const next = this.nextPageUrl(doc, currentUrl, visited);
                if (!next || rows.length === 0) break;
                currentUrl = next;
                await Utils.delay(120);
            }

            return pages;
        },

        contentNearRow(page, row) {
            const rowIndex = page.allRows.indexOf(row);
            if (rowIndex >= 0 && rowIndex < page.allRows.length - 1) {
                const nextRow = page.allRows[rowIndex + 1];
                if (!this.taskId(nextRow)) {
                    const contentCell = nextRow.querySelector('td[colspan] p, td[colspan] div, td[colspan]');
                    if (contentCell) return Utils.cleanRichText(contentCell.innerHTML || contentCell.textContent || '');
                }
            }
            return '';
        }
    };

    // ============================================
    // COLLECTEUR DE TÂCHES TERMINÉES
    // ============================================
    const TasksCompletedCollector = {
        async collect(userId, connectedUser, updateLoader) {
            Utils.log('Collecte des tâches terminées par', connectedUser);
            const results = [];
            const reportDate = Utils.getTodayDate();
            const seen = new Set();

            try {
                const params = new URLSearchParams({
                    'tasks_filters[task_recipient]': userId.taskValue,
                    'tasks_filters[task_status]': 'finished'
                });
                const pages = await TaskListUtils.fetchAll(params, updateLoader, 'Tâches terminées');

                for (const page of pages) {
                    for (const row of page.rows) {
                        const taskId = TaskListUtils.taskId(row);
                        if (!taskId || seen.has(taskId)) continue;

                        const audit = TaskListUtils.audit(row);
                        const fallbackDate = TaskListUtils.dueDate(row).match(/\d{2}\/\d{2}\/\d{4}/)?.[0] || '';
                        const completedDate = audit.lastModifiedDate || fallbackDate;

                        // Une tâche "finished" doit être rattachée au jour où elle a été
                        // clôturée/modifiée, pas à sa date d'échéance.
                        if (completedDate !== reportDate) continue;
                        seen.add(taskId);

                        const titleSpan = row.querySelector('span.font_size_higher');
                        const clientLink = row.querySelector('a[href*="clients_card"]');
                        let content = TaskListUtils.contentNearRow(page, row);

                        if (!content) {
                            if (typeof updateLoader === 'function') updateLoader(`Lecture tâche terminée ${results.length + 1}...`);
                            const taskDetails = await this.getTaskDetails(taskId);
                            content = taskDetails.content || '';
                        }

                        results.push({
                            id: taskId,
                            title: Utils.cleanRichText(titleSpan?.textContent || 'Tâche'),
                            content: Utils.cleanRichText(content),
                            client: Utils.cleanRichText(clientLink?.textContent || 'Non associé'),
                            clientId: clientLink ? (clientLink.href.match(/id=(\d+)/) || [])[1] : null,
                            assignedTo: connectedUser,
                            completedDate,
                            time: audit.lastModifiedTime || '',
                            closedTime: audit.lastModifiedTime || '',
                            closedBy: audit.lastModifiedBy || connectedUser,
                            createdBy: audit.createdBy || '',
                            createdDate: audit.createdDate || '',
                            isPriority: !!row.querySelector('.fa-exclamation'),
                            hasBookmark: !!row.querySelector('.fa-bookmark')
                        });
                    }
                }

                results.sort((a, b) => String(a.closedTime || '').localeCompare(String(b.closedTime || '')));
                Utils.log(`Total: ${results.length} tâches clôturées le ${reportDate}`);
            } catch (error) {
                Utils.log('Erreur collecte tâches terminées:', error);
            }

            return results;
        },

        async getTaskDetails(taskId) {
            try {
                const url = `https://courtage.modulr.fr/fr/scripts/Tasks/TasksCard.php?task_id=${taskId}`;
                const html = await Utils.fetchPage(url);
                const doc = Utils.parseHTML(html);

                const candidates = [
                    doc.querySelector('td[colspan] p'),
                    doc.querySelector('textarea'),
                    doc.querySelector('[name="content"]'),
                    doc.querySelector('[name="description"]')
                ].filter(Boolean);

                let content = '';
                for (const el of candidates) {
                    const value = 'value' in el ? el.value : (el.innerHTML || el.textContent || '');
                    const cleaned = Utils.cleanRichText(value);
                    if (cleaned && cleaned.length > content.length) content = cleaned;
                }
                return { content };
            } catch (error) {
                Utils.log('Erreur lecture tâche:', error);
                return { content: '' };
            }
        }
    };

    // ============================================
    // COLLECTEUR DES TÂCHES ACTUELLEMENT À TRAITER
    // ============================================
    const PendingTasksCollector = {
        async collect(userId, connectedUser, updateLoader) {
            const results = [];
            const seen = new Set();

            try {
                const params = new URLSearchParams({
                    'tasks_filters[task_recipient]': userId.taskValue,
                    'tasks_filters[task_status]': ''
                });
                const pages = await TaskListUtils.fetchAll(params, updateLoader, 'Tâches à traiter');

                for (const page of pages) {
                    for (const row of page.rows) {
                        const taskId = TaskListUtils.taskId(row);
                        if (!taskId || seen.has(taskId)) continue;
                        seen.add(taskId);

                        const title = Utils.cleanRichText(row.querySelector('span.font_size_higher')?.textContent || 'Tâche');
                        const clientLink = row.querySelector('a[href*="clients_card"]');
                        const dueDate = TaskListUtils.dueDate(row);
                        const daysOverdue = TasksOverdueCollector.calculateDaysOverdue(dueDate);
                        const isOverdue = TaskListUtils.isOverdue(row, dueDate);

                        results.push({
                            id: taskId,
                            title,
                            content: TaskListUtils.contentNearRow(page, row),
                            client: Utils.cleanRichText(clientLink?.textContent || 'Non associé'),
                            clientId: clientLink ? (clientLink.href.match(/id=(\d+)/) || [])[1] : null,
                            assignedTo: connectedUser,
                            dueDate,
                            daysOverdue,
                            isOverdue
                        });
                    }
                }

                Utils.log(`${results.length} tâches actuellement à traiter trouvées sur ${pages.length} page(s)`);
            } catch (error) {
                Utils.log('Erreur comptage tâches à traiter:', error);
            }
            return results;
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
                                    content = Utils.cleanRichText(contentCell.innerHTML);
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

            // Depuis les logs : un identifiant de devis/contrat/sinistre n'est PAS
            // un identifiant client. On utilise uniquement client_id lorsqu'il est
            // réellement présent dans les changements du journal.
            [...data.estimates, ...data.policies, ...data.claims, ...data.logs].forEach(log => {
                const clientChange = (log.changes || []).find(change => change.fieldRaw === 'client_id');
                if (!clientChange) return;

                const candidate = [clientChange.newValueRaw, clientChange.oldValueRaw]
                    .find(value => value && value !== '-' && /^\d+$/.test(String(value).trim()));

                if (candidate) {
                    log.clientId = String(candidate).trim();
                    searchItems.push({ type: 'id', value: log.clientId, source: log });
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

            // Logs (devis, contrats, sinistres) : enrichir uniquement avec un
            // véritable client_id identifié dans le journal.
            [...data.estimates, ...data.policies, ...data.claims, ...data.logs].forEach(log => {
                if (log.clientId) {
                    const clientInfo = this.getClientInfo(log.clientId);
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
    // DICTIONNAIRE MÉTIER DES ÉVÉNEMENTS
    // ============================================
    const ActivityDictionary = {
        classify(log) {
            const action = String(log?.actionRaw || log?.action || '').toLowerCase();
            const fields = (log?.changes || []).map(change => change.fieldRaw);

            if (/insertion|création|insert/.test(action)) {
                return { kind: 'creation', label: 'Création', subtype: 'creation' };
            }
            if (/suppression|delete/.test(action)) {
                return { kind: 'deletion', label: 'Suppression', subtype: 'suppression' };
            }

            const rules = [
                { subtype: 'etat', label: 'État', fields: ['status', 'estimate_status', 'policy_status'] },
                { subtype: 'apporteur', label: 'Apporteur', fields: ['producer_id', 'producer', 'business_introducer_id'] },
                { subtype: 'compagnie', label: 'Compagnie', fields: ['company_id', 'company'] },
                { subtype: 'referent', label: 'Référent', fields: ['referent_user_id', 'manager_user_id', 'user_id'] },
                { subtype: 'produit', label: 'Produit', fields: ['product_id', 'product_type_id'] },
                { subtype: 'tarif', label: 'Tarif / prime', fields: ['premium', 'total_amount', 'commission'] },
                { subtype: 'dates', label: 'Dates', fields: ['effective_date', 'start_date', 'end_date', 'renewal_date', 'expiration_date', 'validity_date', 'expiry_date'] },
                { subtype: 'client', label: 'Client', fields: ['client_id'] }
            ];

            for (const rule of rules) {
                if (fields.some(field => rule.fields.includes(field))) {
                    return { kind: 'update', label: rule.label, subtype: rule.subtype };
                }
            }

            return { kind: 'update', label: 'Autre modification', subtype: 'autre' };
        },

        summarize(log) {
            const classification = this.classify(log);
            const meaningful = (log?.changes || []).filter(change =>
                !['last_update', 'last_update_user_id', 'creation_date', 'creation_user_id'].includes(change.fieldRaw)
            );

            if (classification.kind === 'creation') return 'Nouveau dossier créé';
            if (classification.kind === 'deletion') return 'Dossier supprimé';
            if (!meaningful.length) return classification.label;

            const main = meaningful.find(change => {
                if (classification.subtype === 'etat') return change.fieldRaw === 'status';
                if (classification.subtype === 'compagnie') return change.fieldRaw === 'company_id';
                if (classification.subtype === 'apporteur') return /producer/.test(change.fieldRaw);
                if (classification.subtype === 'referent') return /referent|manager_user|^user_id$/.test(change.fieldRaw);
                return true;
            }) || meaningful[0];

            const before = Utils.translateValue(main.oldValueRaw || main.oldValue || '-');
            const after = Utils.translateValue(main.newValueRaw || main.newValue || '-');

            if (before && before !== '-' && after && after !== '-') {
                return `${classification.label} : ${before} → ${after}`;
            }
            if (after && after !== '-') return `${classification.label} → ${after}`;
            return classification.label;
        },

        counts(logs) {
            const result = { creation: 0, update: 0, deletion: 0, subtypes: {} };
            for (const log of logs || []) {
                const c = this.classify(log);
                result[c.kind] = (result[c.kind] || 0) + 1;
                if (c.kind === 'update') {
                    result.subtypes[c.subtype] = (result.subtypes[c.subtype] || 0) + 1;
                }
            }
            return result;
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
            pendingTasks: [],
            logs: [],
            estimates: [],
            policies: [],
            claims: [],
            user: '',
            date: ''
        },

        generateHTML() {
            const {
                emailsSent = [], emailsAffected = [], pendingEmailsCount = 0,
                aircallCalls = [], tasksCompleted = [], tasksOverdue = [], pendingTasks = [],
                logs = [], estimates = [], policies = [], claims = [],
                assignedEstimates = { total: 0, statuses: [] },
                user, date, notes, aircallStatus
            } = this.data;
        
            const realToday = Utils.getRealTodayDate();
            const isPastDate = date !== realToday;
            const estimateCounts = ActivityDictionary.counts(estimates);
            const policyCounts = ActivityDictionary.counts(policies);
            const claimCounts = ActivityDictionary.counts(claims);
            const overduePendingTasks = pendingTasks.filter(task => task.daysOverdue > 0).length;
            const normalPendingTasks = Math.max(0, pendingTasks.length - overduePendingTasks);
            const totalPending = (assignedEstimates.total || 0) + (pendingEmailsCount || 0) + pendingTasks.length;
        
            const subtypeLabels = {
                apporteur: 'Apporteur',
                compagnie: 'Compagnie',
                etat: 'État',
                referent: 'Référent',
                produit: 'Produit',
                tarif: 'Tarif / prime',
                dates: 'Dates',
                client: 'Client',
                autre: 'Autre modification'
            };
        
            const contractTypeLabel = log => {
                const entityName = Utils.cleanRichText(log?.entityName || '');

                const cleanType = value => {
                    let type = String(value || '').trim();

                    // Modulr peut ajouter la compagnie après le type :
                    // "Rapatriement (GRC), - HEOMI"
                    // Dans ce cas HEOMI est la compagnie, pas la typologie.
                    type = type.split(/\s*,\s*-\s*/)[0].trim();

                    // Nettoyage des séparateurs résiduels éventuels.
                    type = type.replace(/[\s,;-]+$/g, '').trim();
                    return type || '—';
                };

                // Le journal fournit le type dans "Nom de la fiche" :
                // "n° 4291 du 30/09/2026 - Automobile"
                if (entityName) {
                    const match = entityName.match(/^\s*n[°ºo]?\s*\d+\s+du\s+\d{1,2}\/\d{1,2}\/\d{4}\s*-\s*(.+?)\s*$/i);
                    if (match && match[1]) return cleanType(match[1]);

                    // Fallback si le préfixe varie légèrement.
                    const prefix = entityName.match(/^.*?\d{1,2}\/\d{1,2}\/\d{4}\s*-\s*(.+)$/);
                    if (prefix && prefix[1]) return cleanType(prefix[1]);
                }

                // Fallback uniquement quand le journal ne fournit vraiment pas le libellé.
                const changes = log?.changes || [];
                const productChange = changes.find(change => change.fieldRaw === 'product_type_id')
                    || changes.find(change => change.fieldRaw === 'product_id');
                if (!productChange) return '—';

                const value = productChange.newValue || productChange.newValueRaw || productChange.oldValue || productChange.oldValueRaw || '';
                return value && value !== '-' ? cleanType(Utils.translateValue(value)) : '—';
            };
        
            const renderLogRows = items => (items || []).map(log => {
                const classification = ActivityDictionary.classify(log);
                const client = log.clientName || (log.clientId ? `Client n° ${log.clientId}` : '—');

                return `
                    <tr data-kind="${classification.kind}" data-subtype="${classification.subtype || ''}">
                        <td class="ltoa-time">${Utils.escapeHtml(log.date || '')}</td>
                        <td><span class="ltoa-tag ltoa-${classification.kind}">${Utils.escapeHtml(classification.label)}</span></td>
                        <td class="ltoa-client">${Utils.escapeHtml(client)}</td>
                        <td class="ltoa-contract-type">${Utils.escapeHtml(contractTypeLabel(log))}</td>
                        <td>${Utils.escapeHtml(ActivityDictionary.summarize(log))}</td>
                    </tr>`;
            }).join('');
        
            const renderMetricButtons = (scope, counts, total, title) => {
                const parts = [];
                if (counts.creation) parts.push(`<button type="button" class="ltoa-metric-btn" data-section="${scope}" data-kind="creation" data-detail-title="${Utils.escapeHtml(title)} · ${counts.creation} création${counts.creation > 1 ? 's' : ''}"><b>${counts.creation}</b> création${counts.creation > 1 ? 's' : ''}</button>`);
                if (counts.update) parts.push(`<button type="button" class="ltoa-metric-btn" data-section="${scope}" data-kind="update" data-detail-title="${Utils.escapeHtml(title)} · ${counts.update} modification${counts.update > 1 ? 's' : ''}"><b>${counts.update}</b> modification${counts.update > 1 ? 's' : ''}</button>`);
                if (counts.deletion) parts.push(`<button type="button" class="ltoa-metric-btn" data-section="${scope}" data-kind="deletion" data-detail-title="${Utils.escapeHtml(title)} · ${counts.deletion} suppression${counts.deletion > 1 ? 's' : ''}"><b>${counts.deletion}</b> suppression${counts.deletion > 1 ? 's' : ''}</button>`);
                return parts.length ? parts.join('') : `<span class="ltoa-muted">Aucune action</span>`;
            };
        
            const renderSubtypeButtons = (scope, counts, title) => {
                const items = Object.entries(counts.subtypes || {}).filter(([, count]) => count > 0);
                if (!items.length) return '';
                return `
                    <div class="ltoa-breakdown">
                        <div class="ltoa-breakdown-label">Détail des modifications</div>
                        <div class="ltoa-breakdown-grid">
                            ${items.map(([key, count]) => `
                                <button type="button" class="ltoa-subtype-btn" data-section="${scope}" data-kind="update" data-subtype="${key}" data-detail-title="${Utils.escapeHtml(title)} · ${count} ${Utils.escapeHtml(subtypeLabels[key] || key)}">
                                    <b>${count}</b>
                                    <span>${Utils.escapeHtml(subtypeLabels[key] || key)}</span>
                                    <small>Voir les dossiers</small>
                                </button>
                            `).join('')}
                        </div>
                    </div>`;
            };


            function renderActionFilters(scope, counts) {
                const total = (counts.creation || 0) + (counts.update || 0) + (counts.deletion || 0);
                return `
                    <div class="ltoa-inline-filters" data-filter-group="${scope}">
                        <span class="ltoa-inline-filter-label">Filtrer par action</span>
                        <button type="button" class="ltoa-inline-filter active" data-filter-kind="">Tous <b>${total}</b></button>
                        ${counts.creation ? `<button type="button" class="ltoa-inline-filter" data-filter-kind="creation">Créations <b>${counts.creation}</b></button>` : ''}
                        ${counts.update ? `<button type="button" class="ltoa-inline-filter" data-filter-kind="update">Modifications <b>${counts.update}</b></button>` : ''}
                        ${counts.deletion ? `<button type="button" class="ltoa-inline-filter" data-filter-kind="deletion">Suppressions <b>${counts.deletion}</b></button>` : ''}
                    </div>`;
            }
        
            const renderTaskCards = items => (items || []).map(task => {
                const text = Utils.cleanRichText(task.content || '');
                return `
                    <div class="ltoa-row-card">
                        <div class="ltoa-row-main">
                            <strong>${Utils.escapeHtml(task.title || 'Tâche')}</strong>
                            <span>${Utils.escapeHtml(task.clientName || task.client || 'Sans client')}</span>
                        </div>
                        <div class="ltoa-row-meta">${Utils.escapeHtml(task.closedTime || task.time || '')}</div>
                        ${text ? `<div class="ltoa-note-text">${Utils.escapeHtml(text)}</div>` : ''}
                    </div>`;
            }).join('');
        
            const section = (id, title, count, summary, body) => `
                <details class="ltoa-section ltoa-section-${id}" id="ltoa-section-${id}">
                    <summary>
                        <div class="ltoa-section-name"><strong>${title}</strong><span class="ltoa-badge">${count}</span></div>
                        <div class="ltoa-section-summary">${summary}</div>
                        <span class="ltoa-chevron">⌄</span>
                    </summary>
                    <div class="ltoa-section-body">${body}</div>
                </details>`;
        
            const table = (headers, rows, emptyText) => rows ? `
                <div class="ltoa-table-wrap">
                    <table class="ltoa-table">
                        <thead><tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr></thead>
                        <tbody>${rows}</tbody>
                    </table>
                </div>` : `<div class="ltoa-empty">${emptyText}</div>`;
        
            const emailRows = [
                ...emailsSent.map(email => `
                    <tr data-kind="sent">
                        <td class="ltoa-time">${Utils.escapeHtml(email.time || email.date || '')}</td>
                        <td><span class="ltoa-tag ltoa-creation">Envoyé</span></td>
                        <td class="ltoa-client">${Utils.escapeHtml(email.clientName || email.toEmail || '—')}</td>
                        <td class="ltoa-ref">Email n° ${Utils.escapeHtml(email.id || '—')}</td>
                        <td>${Utils.escapeHtml(email.subject || 'Sans objet')}</td>
                    </tr>`),
                ...emailsAffected.map(email => `
                    <tr data-kind="affected">
                        <td class="ltoa-time">${Utils.escapeHtml(email.time || email.date || '')}</td>
                        <td><span class="ltoa-tag ltoa-update">Affecté / traité</span></td>
                        <td class="ltoa-client">${Utils.escapeHtml(email.clientName || email.fromEmail || email.from || '—')}</td>
                        <td class="ltoa-ref">Email n° ${Utils.escapeHtml(email.id || '—')}</td>
                        <td>${Utils.escapeHtml(email.subject || 'Sans objet')}</td>
                    </tr>`)
            ].join('');
        
            function renderCallCards() {
                if (!aircallCalls.length) return '<div class="ltoa-empty">Aucun appel</div>';

                const moodMeta = call => {
                    if (call.answered === false || call.missedReason) return { icon: '📵', label: 'Manqué', cls: 'missed' };
                    const mood = String(call.mood || '').toLowerCase();
                    if (mood.includes('posit')) return { icon: '😊', label: 'Positif', cls: 'positive' };
                    if (mood.includes('nég') || mood.includes('neg')) return { icon: '😟', label: 'Négatif', cls: 'negative' };
                    if (mood.includes('neut')) return { icon: '😐', label: 'Neutre', cls: 'neutral' };
                    return { icon: '📞', label: 'Répondu', cls: 'answered' };
                };

                return `
                    <div class="ltoa-call-list">
                        ${aircallCalls.map(call => {
                            const mood = moodMeta(call);
                            const topics = (call.topics || []).map(topic =>
                                typeof topic === 'string' ? topic : (topic?.name || topic?.content || '')
                            ).filter(Boolean);
                            const actions = (call.actionItems || []).filter(Boolean);
                            const kind = call.type === 'entrant' ? 'inbound' : 'outbound';

                            return `
                                <article class="ltoa-call-card ltoa-call-${kind}" data-kind="${kind}">
                                    <div class="ltoa-call-top">
                                        <div>
                                            <div class="ltoa-call-title">
                                                <span class="ltoa-call-direction">${call.type === 'entrant' ? '↙ Entrant' : '↗ Sortant'}</span>
                                                <strong>${Utils.escapeHtml(call.contact || call.phone || 'Contact inconnu')}</strong>
                                            </div>
                                            <div class="ltoa-call-meta">${Utils.escapeHtml(call.time || '')} · ${Utils.escapeHtml(call.duration || '0s')}${call.phone ? ` · ${Utils.escapeHtml(call.phone)}` : ''}</div>
                                        </div>
                                        <div class="ltoa-call-quality ${mood.cls}">
                                            <span>${mood.icon}</span>
                                            <div><small>Qualité / ressenti</small><strong>${Utils.escapeHtml(mood.label)}</strong></div>
                                        </div>
                                    </div>

                                    ${call.summary ? `<div class="ltoa-call-summary"><b>Résumé</b><div>${Utils.escapeHtml(call.summary)}</div></div>` : ''}

                                    ${topics.length ? `
                                        <div class="ltoa-call-detail-block">
                                            <b>Sujets clés</b>
                                            <div class="ltoa-call-pills">${topics.map(topic => `<span>${Utils.escapeHtml(topic)}</span>`).join('')}</div>
                                        </div>` : ''}

                                    ${actions.length ? `
                                        <div class="ltoa-call-detail-block">
                                            <b>Actions à suivre</b>
                                            <div class="ltoa-call-actions">${actions.map(action => `<div>• ${Utils.escapeHtml(action)}</div>`).join('')}</div>
                                        </div>` : ''}

                                    ${call.transcript ? `
                                        <details class="ltoa-call-transcript">
                                            <summary>Voir la transcription${call.transcriptLanguage ? ` · ${Utils.escapeHtml(call.transcriptLanguage)}` : ''}</summary>
                                            <div>${Utils.escapeHtml(call.transcript)}</div>
                                        </details>` : ''}
                                </article>`;
                        }).join('')}
                    </div>`;
            }
        
            const otherRows = logs.map(log => `
                <tr>
                    <td class="ltoa-time">${Utils.escapeHtml(log.date || '')}</td>
                    <td>${Utils.escapeHtml(log.table || log.tableRaw || '')}</td>
                    <td>${Utils.escapeHtml(log.entityName || '')}</td>
                    <td>${Utils.escapeHtml(ActivityDictionary.summarize(log))}</td>
                </tr>`).join('');
        
            const estimateStatusClass = label => {
                const value = String(label || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
                if (value.includes('transform')) return 'is-transformed';
                if (value.includes('perdu') || value.includes('refus')) return 'is-lost';
                if (value.includes('tarification')) return 'is-pricing';
                if (value.includes('souscription')) return 'is-subscription';
                if (value.includes('attente') && value.includes('piece')) return 'is-waiting-docs';
                if (value.includes('attente') && value.includes('appro')) return 'is-waiting-approval';
                if (value.includes('remis') || value.includes('transmis')) return 'is-delivered';
                if (value.includes('diff')) return 'is-deferred';
                if (value.includes('cours')) return 'is-current';
                return 'is-default';
            };

            const assignedStatuses = (assignedEstimates.statuses || []).map(item => `
                <span class="ltoa-work-status ltoa-estimate-status ${estimateStatusClass(item.label)}"><b>${item.count}</b><small>${Utils.escapeHtml(item.label)}</small></span>
            `).join('');
        
            const emailSummary = `
                <button type="button" class="ltoa-metric-btn" data-section="emails" data-kind="sent" data-detail-title="Emails · ${emailsSent.length} envoyés"><b>${emailsSent.length}</b> envoyés</button>
                <button type="button" class="ltoa-metric-btn" data-section="emails" data-kind="affected" data-detail-title="Emails · ${emailsAffected.length} affectés / traités"><b>${emailsAffected.length}</b> affectés / traités</button>`;
        
            const callInbound = aircallCalls.filter(call => call.type === 'entrant').length;
            const callOutbound = aircallCalls.filter(call => call.type === 'sortant').length;
            const callSummary = `
                <button type="button" class="ltoa-metric-btn" data-section="calls" data-kind="inbound" data-detail-title="Appels · ${callInbound} entrants"><b>${callInbound}</b> entrants</button>
                <button type="button" class="ltoa-metric-btn" data-section="calls" data-kind="outbound" data-detail-title="Appels · ${callOutbound} sortants"><b>${callOutbound}</b> sortants</button>`;
        
            return `
                <div id="ltoa-report-modal">
                    <style>
                        #ltoa-report-modal{position:fixed;inset:0;z-index:2147483647;background:radial-gradient(circle at top left,rgba(37,99,235,.07),transparent 27%),#f5f7fb;color:#101828;font-family:Inter,-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;overflow:auto}
                        #ltoa-report-modal *{box-sizing:border-box}
                        .ltoa-shell{min-height:100vh;display:grid;grid-template-columns:minmax(0,1fr)}
                        .ltoa-sidebar{position:sticky;top:0;height:100vh;padding:22px 16px;border-right:1px solid #e7ebf0;background:rgba(255,255,255,.72);backdrop-filter:blur(16px)}
                        .ltoa-brand{display:flex;align-items:center;gap:11px;padding:8px 10px 20px}.ltoa-logo{width:36px;height:36px;border-radius:11px;background:linear-gradient(135deg,#2563eb,#78a8ff);color:#fff;display:grid;place-items:center;font-weight:800;box-shadow:0 10px 22px rgba(37,99,235,.22)}.ltoa-brand strong{font-size:20px}.ltoa-brand small{display:block;color:#7b8795;margin-top:2px}
                        .ltoa-nav{background:rgba(255,255,255,.9);border:1px solid #edf0f4;border-radius:18px;padding:9px;box-shadow:0 12px 30px rgba(15,23,42,.04)}.ltoa-nav-label{font-size:10px;text-transform:uppercase;letter-spacing:.08em;color:#98a2b3;padding:8px 10px}.ltoa-nav-btn{width:100%;border:0;background:transparent;border-radius:12px;padding:10px 11px;text-align:left;cursor:pointer;color:#344054;font-size:12px;margin:2px 0}.ltoa-nav-btn:hover{background:#f5f7fb}.ltoa-nav-btn.active{background:#f7faff;box-shadow:inset 0 0 0 1px #dfe9ff;color:#1d4ed8}
                        .ltoa-content{min-width:0}.ltoa-topbar{position:sticky;top:0;z-index:5;min-height:74px;background:rgba(248,250,253,.88);backdrop-filter:blur(16px);border-bottom:1px solid rgba(228,231,235,.78);display:flex;align-items:center;justify-content:space-between;padding:13px 24px}
                        .ltoa-title h1{font-size:27px;margin:0 0 3px;font-weight:760;letter-spacing:-.03em}.ltoa-title span{font-size:12px;color:#7a8695}.ltoa-actions{display:flex;align-items:center;gap:8px}.ltoa-user-pill{border:1px solid #e5e9ef;background:#fff;border-radius:999px;padding:9px 12px;font-size:12px;font-weight:700;color:#344054}.ltoa-btn{border:1px solid #e1e6ec;background:#fff;border-radius:999px;padding:9px 12px;font-size:12px;cursor:pointer;color:#344150}.ltoa-btn:hover{background:#f7f9fb}.ltoa-close{font-size:18px;line-height:1;padding:7px 10px}
                        .ltoa-main{max-width:1500px;margin:0 auto;padding:22px 24px 42px}
                        .ltoa-warning{background:#fff1f1;border:1px solid #f0c7c7;color:#9b2c2c;border-radius:12px;padding:11px 14px;margin-bottom:14px;font-size:12px}
                .ltoa-report-note{background:linear-gradient(180deg,rgba(255,255,255,.98),rgba(255,255,255,.9));border:1px solid #e4e8ed;border-radius:16px;padding:14px 16px;margin-bottom:14px;box-shadow:0 8px 24px rgba(15,23,42,.025)}.ltoa-report-note strong{display:block;font-size:10px;text-transform:uppercase;letter-spacing:.08em;color:#98a2b3;margin-bottom:6px}.ltoa-report-note div{font-size:13px;line-height:1.55;color:#344054;white-space:pre-wrap}
                        .ltoa-workload{width:100%;background:linear-gradient(180deg,rgba(255,255,255,.96),rgba(255,255,255,.82));border:1px solid rgba(228,231,235,.92);border-radius:24px;padding:20px;box-shadow:0 18px 50px rgba(15,23,42,.06);margin-bottom:24px}
                        .ltoa-work-head{display:flex;justify-content:space-between;align-items:flex-end;gap:20px;margin-bottom:16px}.ltoa-kicker{font-size:10px;color:#98a2b3;text-transform:uppercase;letter-spacing:.09em;margin-bottom:5px}.ltoa-work-head h2,.ltoa-activity-head h2{font-size:21px;margin:0;letter-spacing:-.02em}.ltoa-work-head p,.ltoa-activity-head p{font-size:12px;color:#748091;margin:4px 0 0}.ltoa-work-total{text-align:right}.ltoa-work-total strong{font-size:31px;line-height:1}.ltoa-work-total span{display:block;font-size:10px;color:#7a8695;margin-top:4px}
                        .ltoa-work-grid{display:grid;grid-template-columns:1.45fr .8fr .8fr;gap:10px}.ltoa-work-card{border:1px solid #e7ebf0;background:rgba(255,255,255,.9);border-radius:18px;padding:15px;min-height:126px}.ltoa-work-card.main{border-color:#dbe7ff;background:linear-gradient(180deg,#f8fbff,#fff)}.ltoa-work-card-head{display:flex;justify-content:space-between;gap:12px;align-items:flex-start}.ltoa-work-card-head span{font-size:12px;font-weight:700;color:#344054}.ltoa-work-card-head strong{font-size:26px;line-height:1}.ltoa-work-card p{font-size:10px;color:#7a8695;margin:4px 0 12px}.ltoa-work-statuses{display:flex;flex-wrap:wrap;gap:7px}.ltoa-work-status{display:flex;flex-direction:column;min-width:86px;background:#f7f9fc;border:1px solid #edf0f3;border-radius:12px;padding:8px}.ltoa-work-status b{font-size:14px}.ltoa-work-status small{font-size:9px;color:#748091;margin-top:2px;line-height:1.2}
                        .ltoa-activity-head{margin:4px 0 12px}
                        .ltoa-section{background:rgba(255,255,255,.78);border:1px solid #dfe4ea;border-radius:8px;margin-bottom:8px;overflow:hidden;box-shadow:none}
                        .ltoa-section{--accent:#64748b;--accent-soft:#f1f5f9}.ltoa-section-estimates{--accent:#2563eb;--accent-soft:#eff6ff}.ltoa-section-policies{--accent:#7c3aed;--accent-soft:#f5f3ff}.ltoa-section-emails{--accent:#0891b2;--accent-soft:#ecfeff}.ltoa-section-tasks{--accent:#d97706;--accent-soft:#fffbeb}.ltoa-section-calls{--accent:#ea580c;--accent-soft:#fff7ed}.ltoa-section-claims{--accent:#dc2626;--accent-soft:#fef2f2}.ltoa-section-other{--accent:#475569;--accent-soft:#f8fafc}
                        .ltoa-section>summary{border-left:0}.ltoa-section-name strong{color:var(--accent)}.ltoa-section .ltoa-badge{background:transparent;color:var(--accent);padding:0;font-weight:700}.ltoa-section .ltoa-metric-btn{border-color:#e2e7ec;background:transparent;color:#667085}.ltoa-section .ltoa-metric-btn b{color:var(--accent)}
                        .ltoa-estimate-status.is-current{background:#eff6ff;border-color:#bfdbfe}.ltoa-estimate-status.is-current b{color:#1d4ed8}.ltoa-estimate-status.is-pricing{background:#f5f3ff;border-color:#ddd6fe}.ltoa-estimate-status.is-pricing b{color:#7c3aed}.ltoa-estimate-status.is-delivered{background:#ecfeff;border-color:#a5f3fc}.ltoa-estimate-status.is-delivered b{color:#0e7490}.ltoa-estimate-status.is-waiting-docs{background:#fffbeb;border-color:#fde68a}.ltoa-estimate-status.is-waiting-docs b{color:#b45309}.ltoa-estimate-status.is-waiting-approval{background:#fff7ed;border-color:#fed7aa}.ltoa-estimate-status.is-waiting-approval b{color:#c2410c}.ltoa-estimate-status.is-deferred{background:#f8fafc;border-color:#cbd5e1}.ltoa-estimate-status.is-deferred b{color:#475569}.ltoa-estimate-status.is-lost{background:#fef2f2;border-color:#fecaca}.ltoa-estimate-status.is-lost b{color:#b91c1c}.ltoa-estimate-status.is-transformed{background:#f0fdf4;border-color:#bbf7d0}.ltoa-estimate-status.is-transformed b{color:#15803d}.ltoa-estimate-status.is-subscription{background:#eef2ff;border-color:#c7d2fe}.ltoa-estimate-status.is-subscription b{color:#4338ca}
                        .ltoa-task-status.is-todo{background:#eff6ff;border-color:#bfdbfe}.ltoa-task-status.is-todo b{color:#1d4ed8}.ltoa-task-status.is-overdue{background:#fef2f2;border-color:#fecaca}.ltoa-task-status.is-overdue b{color:#b91c1c}.ltoa-task-status.is-done{background:#f0fdf4;border-color:#bbf7d0}.ltoa-task-status.is-done b{color:#15803d}
                        .ltoa-inline-filters{display:flex;align-items:center;gap:7px;flex-wrap:wrap;margin-bottom:11px}.ltoa-inline-filter-label{font-size:9px;color:#98a2b3;text-transform:uppercase;letter-spacing:.07em;margin-right:2px}.ltoa-inline-filter{border:1px solid #e5e9ef;background:#fff;border-radius:999px;padding:6px 10px;font-size:11px;color:#596579;cursor:pointer}.ltoa-inline-filter b{margin-left:3px}.ltoa-inline-filter.active{background:var(--accent-soft);border-color:var(--accent);color:var(--accent)}
                        .ltoa-call-list{display:block}.ltoa-call-card{border:0;border-bottom:1px solid #e7ebef;background:transparent;border-radius:0;padding:14px 0}.ltoa-call-card:first-child{padding-top:2px}.ltoa-call-card:last-child{border-bottom:0;padding-bottom:2px}.ltoa-call-card.ltoa-call-inbound,.ltoa-call-card.ltoa-call-outbound{border-left:0}.ltoa-call-top{display:flex;align-items:flex-start;justify-content:space-between;gap:14px}.ltoa-call-title{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.ltoa-call-title strong{font-size:14px}.ltoa-call-direction{font-size:11px;font-weight:800;padding:0;background:transparent;color:#475569}.ltoa-call-inbound .ltoa-call-direction{background:transparent;color:#15803d}.ltoa-call-outbound .ltoa-call-direction{background:transparent;color:#1d4ed8}.ltoa-call-meta{font-size:11px;color:#7a8695;margin-top:4px}.ltoa-call-quality{display:flex;align-items:center;gap:6px;border:0;border-radius:0;padding:0;background:transparent;min-width:0}.ltoa-call-quality>span{font-size:18px}.ltoa-call-quality small{display:block;font-size:8px;text-transform:uppercase;color:#98a2b3;letter-spacing:.04em}.ltoa-call-quality strong{display:block;font-size:11px}.ltoa-call-quality.positive,.ltoa-call-quality.negative,.ltoa-call-quality.missed,.ltoa-call-quality.neutral,.ltoa-call-quality.answered{background:transparent;border-color:transparent}.ltoa-call-summary,.ltoa-call-detail-block{margin-top:9px;background:transparent;border-radius:0;padding:0;font-size:12px;line-height:1.5;color:#475569}.ltoa-call-summary b,.ltoa-call-detail-block>b{display:block;font-size:9px;text-transform:uppercase;letter-spacing:.05em;color:#7a8695;margin-bottom:4px}.ltoa-call-pills{display:flex;flex-wrap:wrap;gap:5px}.ltoa-call-pills span{background:transparent;border:0;border-right:1px solid #dfe4ea;border-radius:0;padding:0 8px 0 0;margin-right:2px;font-size:11px}.ltoa-call-pills span:last-child{border-right:0}.ltoa-call-actions div{margin-top:2px}.ltoa-call-transcript{margin-top:9px;border-top:1px solid #edf0f3;padding-top:8px}.ltoa-call-transcript summary{cursor:pointer;color:#c2410c;font-size:10px;font-weight:700}.ltoa-call-transcript div{white-space:pre-wrap;background:transparent;border-left:2px solid #dfe4ea;border-radius:0;padding:4px 0 4px 10px;margin-top:7px;font-size:12px;line-height:1.55;color:#445160}

                        .ltoa-section>summary{list-style:none;cursor:pointer;display:grid;grid-template-columns:minmax(160px,1fr) auto 20px;align-items:center;gap:16px;padding:15px 14px;min-height:58px}.ltoa-section>summary::-webkit-details-marker{display:none}.ltoa-section-name{display:flex;align-items:center;gap:8px}.ltoa-section-name strong{font-size:15px}.ltoa-badge{background:#f0f3f6;border-radius:999px;padding:3px 8px;font-size:10px;color:#536170}.ltoa-section-summary{display:flex;justify-content:flex-end;gap:8px;flex-wrap:wrap}.ltoa-chevron{color:#98a2b3;transition:.18s}.ltoa-section[open] .ltoa-chevron{transform:rotate(180deg)}.ltoa-section[open]>summary{border-bottom:1px solid #edf0f3}.ltoa-section-body{padding:15px 14px 18px}
                        .ltoa-metric-btn{border:1px solid #e5e9ee;background:#fff;border-radius:999px;padding:6px 10px;font-size:11px;color:#5b6674;cursor:pointer}.ltoa-metric-btn:hover{border-color:#cddbf6;background:#f7faff;color:#2456a6}.ltoa-metric-btn b{color:#111827;margin-right:3px}
                        .ltoa-breakdown{background:transparent;border:0;border-top:1px solid #e7ebef;border-bottom:1px solid #e7ebef;border-radius:0;padding:10px 0;margin:10px 0 12px}.ltoa-breakdown-label{font-size:10px;color:#7a8695;text-transform:uppercase;letter-spacing:.06em;margin-bottom:7px}.ltoa-breakdown-grid{display:flex;flex-wrap:wrap;gap:6px}.ltoa-subtype-btn{border:0;border-right:1px solid #e7ebef;background:transparent;border-radius:0;padding:3px 12px 3px 0;margin-right:6px;text-align:left;cursor:pointer;min-width:0}.ltoa-subtype-btn:last-child{border-right:0}.ltoa-subtype-btn:hover{color:var(--accent)}.ltoa-subtype-btn b{display:inline;font-size:13px;margin-right:5px}.ltoa-subtype-btn span{display:inline;font-size:11px;color:#475467;margin:0}.ltoa-subtype-btn small{display:block;font-size:9px;color:#8a94a3;margin-top:2px}
                        .ltoa-table-wrap{overflow:auto;border:1px solid #e5e9ee;border-radius:6px}.ltoa-table{width:100%;border-collapse:collapse;font-size:12px;background:#fff;min-width:760px}.ltoa-table th{background:#f8f9fb;text-align:left;font-weight:700;color:#687584;padding:10px 11px;border-bottom:1px solid #e7ebef;white-space:nowrap;text-transform:uppercase;font-size:10px;letter-spacing:.035em}.ltoa-table td{padding:11px;border-bottom:1px solid #eef1f3;vertical-align:top}.ltoa-table tr:last-child td{border-bottom:0}.ltoa-time{white-space:nowrap;color:#6d7986}.ltoa-client{font-weight:700}.ltoa-contract-type{font-weight:650;color:#344054}.ltoa-tag{display:inline-block;border-radius:999px;padding:3px 7px;font-size:9px;font-weight:700}.ltoa-creation{background:#eaf7ef;color:#267344}.ltoa-update{background:#eef3ff;color:#315ea8}.ltoa-deletion{background:#fff0ef;color:#ae3b32}
                        .ltoa-row-card{border-bottom:1px solid #edf0f2;padding:11px 0}.ltoa-row-card:first-child{padding-top:0}.ltoa-row-card:last-child{border-bottom:0;padding-bottom:0}.ltoa-row-main{display:flex;gap:8px;align-items:baseline;flex-wrap:wrap}.ltoa-row-main strong{font-size:12px}.ltoa-row-main span{font-size:10px;color:#6e7a87}.ltoa-row-meta{font-size:10px;color:#89939d;margin-top:3px}.ltoa-note-text{margin-top:7px;background:#f7f8fa;border-radius:7px;padding:9px 10px;white-space:pre-wrap;line-height:1.45;font-size:11px;color:#3e4954}.ltoa-empty,.ltoa-muted{font-size:11px;color:#8a949e}
                        .ltoa-detail-backdrop{display:none;position:fixed;inset:0;z-index:2147483647;background:rgba(15,23,42,.28);backdrop-filter:blur(6px);align-items:center;justify-content:center;padding:24px}.ltoa-detail-backdrop.show{display:flex}.ltoa-detail-card{width:min(1080px,96vw);max-height:84vh;overflow:auto;background:#fff;border-radius:10px;padding:18px;box-shadow:0 18px 50px rgba(15,23,42,.16)}.ltoa-detail-head{display:flex;align-items:center;justify-content:space-between;gap:14px;margin-bottom:14px}.ltoa-detail-head h3{margin:0;font-size:18px}.ltoa-detail-close{border:0;background:#f1f4f7;border-radius:999px;width:34px;height:34px;cursor:pointer;font-size:18px}
                        @media(max-width:1050px){.ltoa-shell{grid-template-columns:1fr}.ltoa-sidebar{display:none}.ltoa-work-grid{grid-template-columns:1fr}.ltoa-breakdown-grid{grid-template-columns:1fr 1fr}.ltoa-main{padding:16px}.ltoa-topbar{padding:12px 16px}.ltoa-section-summary{justify-content:flex-start}.ltoa-section>summary{grid-template-columns:1fr 20px}.ltoa-section-summary{display:none}}
                    </style>
        
                    <div class="ltoa-shell">
                        <div class="ltoa-content">
                            <header class="ltoa-topbar" id="ltoa-report-top">
                                <div class="ltoa-title">
                                    <h1>Rapport d’activité</h1>
                                    <span>${Utils.escapeHtml(date || '')}${isPastDate ? ' · rétrospectif' : ''}</span>
                                </div>
                                <div class="ltoa-actions">
                                    <span class="ltoa-user-pill">${Utils.escapeHtml(user || '')}</span>
                                    <button id="ltoa-view-by-client" class="ltoa-btn">Par client</button>
                                    <button id="ltoa-view-chrono" class="ltoa-btn">Chronologie</button>
                                    <button id="ltoa-export-html" class="ltoa-btn">Exporter</button>
                                    <button id="ltoa-close-report" class="ltoa-btn ltoa-close">×</button>
                                </div>
                            </header>
        
                            <main class="ltoa-main">
                                ${aircallStatus?.state === 'error' ? `<div class="ltoa-warning">${Utils.escapeHtml(aircallStatus.message || 'Erreur Aircall')}</div>` : ''}
                                ${notes ? `<div class="ltoa-report-note"><strong>Note</strong><div>${Utils.escapeHtml(notes)}</div></div>` : ''}
        
                                <section class="ltoa-workload">
                                    <div class="ltoa-work-head">
                                        <div>
                                            <div class="ltoa-kicker">Travail à faire</div>
                                            <h2>Charge actuelle</h2>
                                            <p>Les éléments actuellement affectés à ${Utils.escapeHtml(user || 'ce collaborateur')} et encore à traiter.</p>
                                        </div>
                                        <div class="ltoa-work-total"><strong>${totalPending}</strong><span>éléments à traiter</span></div>
                                    </div>
                                    <div class="ltoa-work-grid">
                                        <div class="ltoa-work-card main">
                                            <div class="ltoa-work-card-head"><span>Devis à traiter</span><strong>${assignedEstimates.total || 0}</strong></div>
                                            <p>Devis actuellement assignés</p>
                                            <div class="ltoa-work-statuses">${assignedStatuses || '<span class="ltoa-muted">Aucun devis assigné</span>'}</div>
                                        </div>
                                        <div class="ltoa-work-card">
                                            <div class="ltoa-work-card-head"><span>Emails à traiter</span><strong>${pendingEmailsCount || 0}</strong></div>
                                            <p>Emails affectés toujours en attente de traitement</p>
                                        </div>
                                        <div class="ltoa-work-card">
                                            <div class="ltoa-work-card-head"><span>Tâches à traiter</span><strong>${pendingTasks.length}</strong></div>
                                            <p>Tâches ouvertes affectées au collaborateur</p>
                                            <div class="ltoa-work-statuses">
                                                <span class="ltoa-work-status ltoa-task-status is-todo"><b>${normalPendingTasks}</b><small>À faire</small></span>
                                                <span class="ltoa-work-status ltoa-task-status is-overdue"><b>${overduePendingTasks}</b><small>En retard</small></span>
                                                <span class="ltoa-work-status ltoa-task-status is-done"><b>${tasksCompleted.length}</b><small>Traitées aujourd’hui</small></span>
                                            </div>
                                        </div>
                                    </div>
                                </section>
        
                                <div class="ltoa-activity-head">
                                    <div class="ltoa-kicker">Travail effectué</div>
                                    <h2>Activité du jour</h2>
                                    <p>Tous les modules sont repliés. Cliquez sur un volume pour afficher uniquement les dossiers correspondants.</p>
                                </div>
        
                                ${section(
                                    'estimates',
                                    'Devis',
                                    estimates.length,
                                    renderMetricButtons('estimates', estimateCounts, estimates.length, 'Devis'),
                                    `
                                        ${renderActionFilters('estimates', estimateCounts)}
                                        ${renderSubtypeButtons('estimates', estimateCounts, 'Devis')}
                                        ${table(['Date','Action','Client','Type de contrat','Détail'], renderLogRows(estimates), 'Aucune action sur les devis')}
                                    `
                                )}
        
                                ${section(
                                    'policies',
                                    'Contrats',
                                    policies.length,
                                    renderMetricButtons('policies', policyCounts, policies.length, 'Contrats'),
                                    `
                                        ${renderActionFilters('policies', policyCounts)}
                                        ${renderSubtypeButtons('policies', policyCounts, 'Contrats')}
                                        ${table(['Date','Action','Client','Type de contrat','Détail'], renderLogRows(policies), 'Aucune action sur les contrats')}
                                    `
                                )}
        
                                ${section(
                                    'emails',
                                    'Emails',
                                    emailsSent.length + emailsAffected.length,
                                    emailSummary,
                                    table(['Heure','Action','Client / interlocuteur','Référence','Objet'], emailRows, 'Aucune activité email')
                                )}
        
                                ${section(
                                    'tasks',
                                    'Tâches',
                                    tasksCompleted.length,
                                    `<span class="ltoa-metric-btn" style="cursor:default"><b>${tasksCompleted.length}</b> terminées</span>`,
                                    tasksCompleted.length ? renderTaskCards(tasksCompleted) : '<div class="ltoa-empty">Aucune tâche terminée aujourd’hui</div>'
                                )}
        
                                ${section(
                                    'calls',
                                    'Appels',
                                    aircallCalls.length,
                                    callSummary,
                                    renderCallCards()
                                )}
        
                                ${section(
                                    'claims',
                                    'Sinistres',
                                    claims.length,
                                    renderMetricButtons('claims', claimCounts, claims.length, 'Sinistres'),
                                    `
                                        ${renderSubtypeButtons('claims', claimCounts, 'Sinistres')}
                                        ${table(['Date','Action','Client','Type','Détail'], renderLogRows(claims), 'Aucune action sur les sinistres')}
                                    `
                                )}
        
                                ${logs.length ? section(
                                    'other',
                                    'Autres actions',
                                    logs.length,
                                    `<span class="ltoa-muted">${logs.length} action${logs.length > 1 ? 's' : ''} complémentaire${logs.length > 1 ? 's' : ''}</span>`,
                                    table(['Date','Rubrique','Élément','Détail'], otherRows, 'Aucune autre action utile')
                                ) : ''}
                            </main>
                        </div>
                    </div>
        
                    <div class="ltoa-detail-backdrop" id="ltoa-detail-backdrop">
                        <div class="ltoa-detail-card">
                            <div class="ltoa-detail-head">
                                <h3 id="ltoa-detail-title">Détail</h3>
                                <button type="button" class="ltoa-detail-close" id="ltoa-detail-close">×</button>
                            </div>
                            <div id="ltoa-detail-content"></div>
                        </div>
                    </div>
                </div>`;
        },

        show() {
            const existing = document.getElementById('ltoa-report-modal');
            if (existing) existing.remove();
        
            document.body.insertAdjacentHTML('beforeend', this.generateHTML());
            const modal = document.getElementById('ltoa-report-modal');
            const detailBackdrop = document.getElementById('ltoa-detail-backdrop');
            const detailTitle = document.getElementById('ltoa-detail-title');
            const detailContent = document.getElementById('ltoa-detail-content');
        
            document.getElementById('ltoa-close-report').addEventListener('click', () => modal.remove());
            document.getElementById('ltoa-view-by-client').addEventListener('click', () => this.showByClientView());
            document.getElementById('ltoa-export-html').addEventListener('click', () => this.exportHTML());
            document.getElementById('ltoa-view-chrono').addEventListener('click', () => this.showChronoView());
        
            const closeDetail = () => detailBackdrop.classList.remove('show');
            document.getElementById('ltoa-detail-close').addEventListener('click', closeDetail);
            detailBackdrop.addEventListener('click', event => {
                if (event.target === detailBackdrop) closeDetail();
            });
        
            const openMetricDetail = button => {
                const sectionId = button.dataset.section;
                const kind = button.dataset.kind || '';
                const subtype = button.dataset.subtype || '';
                const sectionEl = document.getElementById(`ltoa-section-${sectionId}`);
                const sourceTable = sectionEl?.querySelector('.ltoa-table');
                const sourceCalls = sectionEl?.querySelector('.ltoa-call-list');
                if (!sourceTable && !sourceCalls) return;

                detailTitle.textContent = button.dataset.detailTitle || 'Détail';
                detailContent.innerHTML = '';

                if (sourceTable) {
                    const clone = sourceTable.cloneNode(true);
                    const rows = Array.from(clone.querySelectorAll('tbody tr'));
                    rows.forEach(row => {
                        const matchesKind = !kind || row.dataset.kind === kind;
                        const matchesSubtype = !subtype || row.dataset.subtype === subtype;
                        if (!matchesKind || !matchesSubtype) row.remove();
                    });
                    const wrap = document.createElement('div');
                    wrap.className = 'ltoa-table-wrap';
                    wrap.appendChild(clone);
                    detailContent.appendChild(wrap);
                } else if (sourceCalls) {
                    const clone = sourceCalls.cloneNode(true);
                    clone.querySelectorAll('.ltoa-call-card').forEach(card => {
                        if (kind && card.dataset.kind !== kind) card.remove();
                    });
                    detailContent.appendChild(clone);
                }

                detailBackdrop.classList.add('show');
            };
        
            modal.querySelectorAll('.ltoa-metric-btn[data-section], .ltoa-subtype-btn[data-section]').forEach(button => {
                button.addEventListener('click', event => {
                    event.preventDefault();
                    event.stopPropagation();
                    openMetricDetail(button);
                });
            });
        
            modal.querySelectorAll('.ltoa-inline-filters').forEach(group => {
                const sectionEl = group.closest('.ltoa-section');
                const rows = sectionEl ? Array.from(sectionEl.querySelectorAll('.ltoa-table tbody tr')) : [];

                group.querySelectorAll('.ltoa-inline-filter').forEach(button => {
                    button.addEventListener('click', event => {
                        event.preventDefault();
                        event.stopPropagation();

                        const kind = button.dataset.filterKind || '';
                        group.querySelectorAll('.ltoa-inline-filter').forEach(item => item.classList.remove('active'));
                        button.classList.add('active');

                        rows.forEach(row => {
                            row.style.display = !kind || row.dataset.kind === kind ? '' : 'none';
                        });
                    });
                });
            });
        
            document.addEventListener('keydown', event => {
                if (event.key === 'Escape') {
                    if (detailBackdrop?.classList.contains('show')) {
                        closeDetail();
                    } else {
                        const current = document.getElementById('ltoa-report-modal');
                        if (current) current.remove();
                    }
                }
            }, { once: false });
        },

        // ============================================
        // VUE PAR CLIENT
        // ============================================
        showByClientView() {
    try {
        const {
            emailsSent = [], emailsAffected = [], aircallCalls = [],
            tasksCompleted = [], tasksOverdue = [], logs = [],
            estimates = [], policies = [], claims = [],
            user = '', date = ''
        } = this.data;

        const existing = document.getElementById('ltoa-client-view-modal');
        if (existing) existing.remove();

        const normalize = value => String(value || '')
            .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
            .toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();

        const clientsMap = new Map();

        const addToClient = (clientName, clientId, clientEmail, type, item) => {
            const cleanName = clientName && !['N/A', 'Non associé', 'Non associe'].includes(clientName)
                ? String(clientName).trim()
                : '';
            const key = clientId ? `id_${clientId}` : (cleanName ? `name_${normalize(cleanName)}` : 'unresolved');

            if (!clientsMap.has(key)) {
                clientsMap.set(key, {
                    name: cleanName || 'Sans client associé',
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

            const client = clientsMap.get(key);
            if (cleanName && client.name === 'Sans client associé') client.name = cleanName;
            if (clientId && !client.id) client.id = clientId;
            if (clientEmail && !client.email) client.email = clientEmail;
            if (client[type]) client[type].push(item);
        };

        emailsSent.forEach(item => addToClient(
            item.clientName || item.recipientName || item.to || item.toEmail,
            item.clientId,
            item.clientEmail || item.toEmail,
            'emailsSent',
            item
        ));

        emailsAffected.forEach(item => addToClient(
            item.clientName || item.fromName || item.from || item.fromEmail,
            item.clientId,
            item.clientEmail || item.fromEmail,
            'emailsAffected',
            item
        ));

        tasksCompleted.forEach(item => addToClient(
            item.clientName || item.client,
            item.clientId,
            item.clientEmail,
            'tasksCompleted',
            item
        ));

        tasksOverdue.forEach(item => addToClient(
            item.clientName || item.client,
            item.clientId,
            item.clientEmail,
            'tasksOverdue',
            item
        ));

        estimates.forEach(item => addToClient(
            item.clientName || item.entityName,
            item.clientId,
            item.clientEmail,
            'estimates',
            item
        ));

        policies.forEach(item => addToClient(
            item.clientName || item.entityName,
            item.clientId,
            item.clientEmail,
            'policies',
            item
        ));

        claims.forEach(item => addToClient(
            item.clientName || item.entityName,
            item.clientId,
            item.clientEmail,
            'claims',
            item
        ));

        logs.forEach(item => {
            const table = normalize(`${item.tableRaw || ''} ${item.table || ''}`);
            if (/utilisateur|user|collaborateur|employe/.test(table)) return;

            let clientId = item.clientId || null;
            let clientName = item.clientName || null;

            if (!clientId && /\bclient/.test(table) && /^\d+$/.test(String(item.entityId || ''))) {
                clientId = String(item.entityId);
                clientName = item.entityName || clientName;
            }

            if (!clientId && Array.isArray(item.changes)) {
                const change = item.changes.find(c => c.fieldRaw === 'client_id');
                const candidate = change
                    ? [change.newValueRaw, change.oldValueRaw].find(value => /^\d+$/.test(String(value || '').trim()))
                    : null;
                if (candidate) clientId = String(candidate).trim();
            }

            if (clientId || clientName) {
                addToClient(clientName || item.entityName, clientId, item.clientEmail, 'logs', item);
            }
        });

        // Les appels Aircall n'ont pas toujours d'identifiant Modulr.
        // On les rattache seulement lorsqu'un nom correspond suffisamment,
        // sinon ils restent dans une fiche "contact" séparée.
        aircallCalls.forEach(call => {
            const contactName = call.clientName || call.contact || call.phone || 'Contact inconnu';
            const contactNorm = normalize(contactName);
            const contactParts = contactNorm.split(' ').filter(part => part.length > 2);
            let best = null;
            let bestScore = 0;

            clientsMap.forEach(client => {
                if (!client.name || client.name === 'Sans client associé') return;
                const clientNorm = normalize(client.name);
                if (!clientNorm) return;

                if (clientNorm === contactNorm) {
                    best = client;
                    bestScore = 100;
                    return;
                }

                const clientParts = clientNorm.split(' ').filter(part => part.length > 2);
                const score = contactParts.filter(part =>
                    clientParts.some(other => other === part || other.includes(part) || part.includes(other))
                ).length;

                if (score > bestScore) {
                    best = client;
                    bestScore = score;
                }
            });

            if (best && (bestScore >= 2 || (contactParts.length === 1 && bestScore === 1))) {
                best.aircallCalls.push(call);
            } else {
                addToClient(contactName, call.clientId, null, 'aircallCalls', call);
            }
        });

        const mergeById = new Map();
        clientsMap.forEach(client => {
            const key = client.id ? `id_${client.id}` : `name_${normalize(client.name) || 'unresolved'}`;
            if (!mergeById.has(key)) {
                mergeById.set(key, client);
                return;
            }
            const target = mergeById.get(key);
            for (const field of ['emailsSent','emailsAffected','aircallCalls','tasksCompleted','tasksOverdue','estimates','policies','claims','logs']) {
                target[field].push(...client[field]);
            }
            if (!target.email && client.email) target.email = client.email;
            if ((!target.name || target.name === 'Sans client associé') && client.name) target.name = client.name;
        });

        const activityCount = client =>
            client.emailsSent.length + client.emailsAffected.length + client.aircallCalls.length +
            client.tasksCompleted.length + client.estimates.length + client.policies.length +
            client.claims.length + client.logs.length;

        const activeClients = Array.from(mergeById.values())
            .filter(client => activityCount(client) > 0)
            .sort((a, b) => activityCount(b) - activityCount(a) || String(a.name).localeCompare(String(b.name), 'fr'));

        const renderText = text => {
            const clean = Utils.cleanRichText(text || '');
            if (!clean) return '';
            if (clean.length <= 220) return `<div class="ltoa-sub-text">${Utils.escapeHtml(clean)}</div>`;
            return `
                <details class="ltoa-more">
                    <summary>Voir le contenu complet</summary>
                    <div class="ltoa-sub-text">${Utils.escapeHtml(clean)}</div>
                </details>`;
        };

        const renderGroup = (title, count, body) => {
            if (!count) return '';
            return `
                <section class="ltoa-client-group">
                    <div class="ltoa-client-group-head">
                        <strong>${Utils.escapeHtml(title)}</strong>
                        <span>${count}</span>
                    </div>
                    <div class="ltoa-client-group-body">${body}</div>
                </section>`;
        };

        const renderDocument = (item, type) => {
            const classification = ActivityDictionary.classify(item);
            return `
                <div class="ltoa-sub-row">
                    <div>
                        <strong>${Utils.escapeHtml(type)} n° ${Utils.escapeHtml(item.entityId || '—')}</strong>
                        <small>${Utils.escapeHtml(classification.label)}</small>
                    </div>
                    <div class="ltoa-sub-row-detail">${Utils.escapeHtml(ActivityDictionary.summarize(item))}</div>
                </div>`;
        };

        const cards = activeClients.map((client, index) => {
            const total = activityCount(client);
            const clientLink = client.id
                ? `https://courtage.modulr.fr/fr/scripts/clients/clients_card.php?id=${encodeURIComponent(client.id)}`
                : '';
            const search = normalize(`${client.name} ${client.email || ''} ${client.id || ''}`);

            const sent = client.emailsSent.map(email => `
                <div class="ltoa-sub-row">
                    <div>
                        <strong>${Utils.escapeHtml(email.subject || 'Sans objet')}</strong>
                        <small>${Utils.escapeHtml(email.time || email.date || '')} · À ${Utils.escapeHtml(email.toEmail || email.to || '—')}</small>
                    </div>
                    ${renderText(email.body)}
                </div>`).join('');

            const received = client.emailsAffected.map(email => `
                <div class="ltoa-sub-row">
                    <div>
                        <strong>${Utils.escapeHtml(email.subject || 'Sans objet')}</strong>
                        <small>${Utils.escapeHtml(email.time || email.date || '')} · De ${Utils.escapeHtml(email.from || email.fromEmail || '—')}</small>
                    </div>
                    ${email.affectedTo ? `<div class="ltoa-sub-row-detail">Affecté à ${Utils.escapeHtml(email.affectedTo)}</div>` : ''}
                    ${renderText(email.body)}
                </div>`).join('');

            const calls = client.aircallCalls.map(call => `
                <div class="ltoa-sub-row">
                    <div class="ltoa-call-head">
                        <div>
                            <strong>${call.type === 'sortant' ? 'Appel sortant' : 'Appel entrant'} · ${Utils.escapeHtml(call.contact || call.phone || '—')}</strong>
                            <small>${Utils.escapeHtml(call.time || '')} · ${Utils.escapeHtml(call.duration || '0s')}${call.mood ? ` · ${Utils.escapeHtml(call.mood)}` : ''}</small>
                        </div>
                    </div>
                    ${call.summary ? `<div class="ltoa-sub-text"><b>Résumé</b><br>${Utils.escapeHtml(call.summary)}</div>` : ''}
                    ${call.topics?.length ? `<div class="ltoa-mini-pills">${call.topics.map(topic => `<span>${Utils.escapeHtml(typeof topic === 'string' ? topic : (topic?.name || topic?.content || ''))}</span>`).join('')}</div>` : ''}
                    ${call.actionItems?.length ? `
                        <div class="ltoa-sub-text"><b>Actions à suivre</b><br>${call.actionItems.map(item => `• ${Utils.escapeHtml(item)}`).join('<br>')}</div>` : ''}
                    ${call.transcript ? `
                        <details class="ltoa-more">
                            <summary>Voir la transcription</summary>
                            <div class="ltoa-sub-text">${Utils.escapeHtml(call.transcript)}</div>
                        </details>` : ''}
                </div>`).join('');

            const completedTasks = client.tasksCompleted.map(task => `
                <div class="ltoa-sub-row">
                    <div>
                        <strong>${Utils.escapeHtml(task.title || 'Tâche')}</strong>
                        <small>${Utils.escapeHtml(task.closedTime || task.time || task.completedDate || '')}${task.closedBy ? ` · clôturée par ${Utils.escapeHtml(task.closedBy)}` : ''}</small>
                    </div>
                    ${renderText(task.content)}
                </div>`).join('');

            const overdueTasks = client.tasksOverdue.map(task => `
                <div class="ltoa-sub-row">
                    <div>
                        <strong>${Utils.escapeHtml(task.title || 'Tâche')}</strong>
                        <small>${Utils.escapeHtml(task.dueDate || '')}${task.daysOverdue ? ` · ${task.daysOverdue} j de retard` : ''}</small>
                    </div>
                    ${renderText(task.content)}
                </div>`).join('');

            const documents = [
                ...client.estimates.map(item => renderDocument(item, 'Devis')),
                ...client.policies.map(item => renderDocument(item, 'Contrat')),
                ...client.claims.map(item => renderDocument(item, 'Sinistre'))
            ].join('');

            const clientLogs = client.logs.map(log => `
                <div class="ltoa-sub-row">
                    <div>
                        <strong>${Utils.escapeHtml(log.table || log.tableRaw || 'Modification')}</strong>
                        <small>${Utils.escapeHtml(log.date || '')} · ${Utils.escapeHtml(log.action || log.actionRaw || '')}</small>
                    </div>
                    <div class="ltoa-sub-row-detail">${Utils.escapeHtml(ActivityDictionary.summarize(log))}</div>
                </div>`).join('');

            return `
                <details class="ltoa-client-card" data-search="${Utils.escapeHtml(search)}" ${index < 3 ? 'open' : ''}>
                    <summary>
                        <div class="ltoa-client-ident">
                            <strong>${Utils.escapeHtml(client.name || 'Sans client associé')}</strong>
                            <span>${client.id ? `Client n° ${Utils.escapeHtml(client.id)}` : 'Contact non relié à une fiche Modulr'}${client.email ? ` · ${Utils.escapeHtml(client.email)}` : ''}</span>
                        </div>
                        <div class="ltoa-client-counts">
                            ${client.emailsSent.length ? `<span>${client.emailsSent.length} envoyés</span>` : ''}
                            ${client.emailsAffected.length ? `<span>${client.emailsAffected.length} reçus</span>` : ''}
                            ${client.aircallCalls.length ? `<span>${client.aircallCalls.length} appels</span>` : ''}
                            ${client.tasksCompleted.length ? `<span>${client.tasksCompleted.length} tâches</span>` : ''}
                            ${(client.estimates.length + client.policies.length + client.claims.length) ? `<span>${client.estimates.length + client.policies.length + client.claims.length} dossiers</span>` : ''}
                            <b>${total}</b>
                        </div>
                    </summary>
                    <div class="ltoa-client-body">
                        ${clientLink ? `<a class="ltoa-client-link" href="${clientLink}" target="_blank" rel="noopener">Ouvrir la fiche client dans Modulr ↗</a>` : ''}
                        ${renderGroup('Emails envoyés', client.emailsSent.length, sent)}
                        ${renderGroup('Emails reçus / affectés', client.emailsAffected.length, received)}
                        ${renderGroup('Appels', client.aircallCalls.length, calls)}
                        ${renderGroup('Tâches terminées', client.tasksCompleted.length, completedTasks)}
                        ${renderGroup('Tâches en retard', client.tasksOverdue.length, overdueTasks)}
                        ${renderGroup('Devis, contrats et sinistres', client.estimates.length + client.policies.length + client.claims.length, documents)}
                        ${renderGroup('Autres modifications de la fiche', client.logs.length, clientLogs)}
                    </div>
                </details>`;
        }).join('');

        const totalActions = activeClients.reduce((sum, client) => sum + activityCount(client), 0);

        const html = `
            <div id="ltoa-client-view-modal" class="ltoa-secondary-modal">
                <style>
                    #ltoa-client-view-modal{position:fixed;inset:0;z-index:2147483647;background:#f5f7fb;color:#101828;font-family:Inter,-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;overflow:auto}
                    #ltoa-client-view-modal *{box-sizing:border-box}
                    .ltoa-secondary-top{position:sticky;top:0;z-index:10;background:rgba(248,250,253,.92);backdrop-filter:blur(16px);border-bottom:1px solid #e5e9ef;padding:13px 22px;display:flex;align-items:center;justify-content:space-between;gap:16px}
                    .ltoa-secondary-title h2{margin:0;font-size:25px;letter-spacing:-.03em}.ltoa-secondary-title p{margin:4px 0 0;color:#7a8695;font-size:12px}
                    .ltoa-secondary-actions{display:flex;align-items:center;gap:8px}.ltoa-secondary-btn{border:1px solid #dfe4ea;background:#fff;border-radius:999px;padding:9px 12px;font-size:12px;color:#344054;cursor:pointer}.ltoa-secondary-btn:hover{background:#f7f9fb}
                    .ltoa-secondary-main{max-width:1280px;margin:0 auto;padding:22px 24px 44px}
                    .ltoa-secondary-summary{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;margin-bottom:14px}.ltoa-summary-card{background:#fff;border:1px solid #e4e8ed;border-radius:17px;padding:14px 16px}.ltoa-summary-card strong{display:block;font-size:24px}.ltoa-summary-card span{font-size:10px;color:#7a8695}
                    .ltoa-client-tools{display:flex;gap:8px;align-items:center;margin:0 0 14px}.ltoa-client-search{flex:1;border:1px solid #dfe4ea;background:#fff;border-radius:14px;padding:11px 13px;font:inherit;font-size:12px;outline:none}.ltoa-client-search:focus{border-color:#a9c0f7;box-shadow:0 0 0 4px rgba(37,99,235,.07)}
                    .ltoa-client-card{background:#fff;border:1px solid #e4e8ed;border-radius:18px;margin-bottom:10px;overflow:hidden;box-shadow:0 8px 24px rgba(15,23,42,.025)}.ltoa-client-card>summary{list-style:none;cursor:pointer;padding:15px 16px;display:grid;grid-template-columns:minmax(0,1fr) auto;gap:14px;align-items:center}.ltoa-client-card>summary::-webkit-details-marker{display:none}.ltoa-client-card[open]>summary{border-bottom:1px solid #edf0f3;background:#fbfcfe}
                    .ltoa-client-ident strong{display:block;font-size:14px}.ltoa-client-ident span{display:block;color:#7a8695;font-size:10px;margin-top:3px}.ltoa-client-counts{display:flex;gap:6px;align-items:center;justify-content:flex-end;flex-wrap:wrap}.ltoa-client-counts span{font-size:9px;border:1px solid #e6eaf0;background:#f8fafc;border-radius:999px;padding:5px 8px;color:#596579}.ltoa-client-counts b{min-width:31px;text-align:center;border-radius:999px;background:#eef3ff;color:#315ea8;padding:6px 8px;font-size:11px}
                    .ltoa-client-body{padding:14px 16px 17px}.ltoa-client-link{display:inline-block;margin-bottom:12px;color:#2259da;text-decoration:none;font-size:11px;font-weight:700}.ltoa-client-group{border-top:1px solid #edf0f3;padding:13px 0}.ltoa-client-group:first-of-type{border-top:0}.ltoa-client-group-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:7px}.ltoa-client-group-head strong{font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:#315ea8}.ltoa-client-group:nth-of-type(2) .ltoa-client-group-head strong{color:#0891b2}.ltoa-client-group:nth-of-type(3) .ltoa-client-group-head strong{color:#ea580c}.ltoa-client-group:nth-of-type(4) .ltoa-client-group-head strong{color:#d97706}.ltoa-client-group:nth-of-type(5) .ltoa-client-group-head strong{color:#dc2626}.ltoa-client-group:nth-of-type(6) .ltoa-client-group-head strong{color:#7c3aed}.ltoa-client-group-head span{font-size:10px;background:#f1f4f7;border-radius:999px;padding:3px 7px;color:#667085}
                    .ltoa-client-group-body{display:grid;gap:7px}.ltoa-sub-row{border:1px solid #edf0f3;background:#fcfdff;border-radius:13px;padding:10px 11px}.ltoa-sub-row strong{font-size:11px}.ltoa-sub-row small{display:block;margin-top:3px;font-size:9px;color:#84909f}.ltoa-sub-row-detail{font-size:10px;color:#526071;margin-top:6px}.ltoa-sub-text{white-space:pre-wrap;background:#f5f7fa;border-radius:9px;padding:8px 9px;margin-top:7px;font-size:10px;line-height:1.45;color:#445160}.ltoa-more{margin-top:7px}.ltoa-more summary{cursor:pointer;color:#315ea8;font-size:10px;font-weight:700}.ltoa-mini-pills{display:flex;gap:5px;flex-wrap:wrap;margin-top:7px}.ltoa-mini-pills span{font-size:9px;border:1px solid #e6eaf0;border-radius:999px;padding:4px 7px;background:#fff;color:#667085}
                    .ltoa-empty-secondary{background:#fff;border:1px solid #e4e8ed;border-radius:18px;padding:42px;text-align:center;color:#7a8695;font-size:12px}
                    @media(max-width:760px){.ltoa-secondary-top{align-items:flex-start}.ltoa-secondary-actions{flex-wrap:wrap;justify-content:flex-end}.ltoa-secondary-main{padding:14px}.ltoa-secondary-summary{grid-template-columns:1fr}.ltoa-client-tools{flex-wrap:wrap}.ltoa-client-search{flex-basis:100%}.ltoa-client-card>summary{grid-template-columns:1fr}.ltoa-client-counts{justify-content:flex-start}}
                </style>
                <header class="ltoa-secondary-top">
                    <div class="ltoa-secondary-title">
                        <h2>Vue par client</h2>
                        <p>${Utils.escapeHtml(user)} · ${Utils.escapeHtml(date)} · même activité, regroupée dossier par dossier</p>
                    </div>
                    <div class="ltoa-secondary-actions">
                        <button type="button" class="ltoa-secondary-btn" id="ltoa-client-expand">Tout ouvrir</button>
                        <button type="button" class="ltoa-secondary-btn" id="ltoa-client-collapse">Tout replier</button>
                        <button type="button" class="ltoa-secondary-btn" id="ltoa-close-client-view">Fermer</button>
                    </div>
                </header>
                <main class="ltoa-secondary-main">
                    <div class="ltoa-secondary-summary">
                        <div class="ltoa-summary-card"><strong>${activeClients.length}</strong><span>clients / contacts concernés</span></div>
                        <div class="ltoa-summary-card"><strong>${totalActions}</strong><span>actions rattachées à un client</span></div>
                        <div class="ltoa-summary-card"><strong>${tasksCompleted.length}</strong><span>tâches terminées sur la journée</span></div>
                    </div>
                    <div class="ltoa-client-tools">
                        <input id="ltoa-client-search" class="ltoa-client-search" type="search" placeholder="Rechercher un client, un email ou un numéro de fiche…">
                    </div>
                    <div id="ltoa-client-list">
                        ${cards || '<div class="ltoa-empty-secondary">Aucun client ou contact concerné par l’activité de cette journée.</div>'}
                    </div>
                </main>
            </div>`;

        document.body.insertAdjacentHTML('beforeend', html);
        const modal = document.getElementById('ltoa-client-view-modal');
        const close = () => modal?.remove();

        document.getElementById('ltoa-close-client-view')?.addEventListener('click', close);
        document.getElementById('ltoa-client-expand')?.addEventListener('click', () => {
            modal.querySelectorAll('.ltoa-client-card').forEach(card => card.open = true);
        });
        document.getElementById('ltoa-client-collapse')?.addEventListener('click', () => {
            modal.querySelectorAll('.ltoa-client-card').forEach(card => card.open = false);
        });
        document.getElementById('ltoa-client-search')?.addEventListener('input', event => {
            const query = normalize(event.target.value);
            modal.querySelectorAll('.ltoa-client-card').forEach(card => {
                const match = !query || String(card.dataset.search || '').includes(query);
                card.style.display = match ? '' : 'none';
                if (query && match) card.open = true;
            });
        });

        const escHandler = event => {
            if (event.key === 'Escape' && document.getElementById('ltoa-client-view-modal')) {
                close();
                document.removeEventListener('keydown', escHandler);
            }
        };
        document.addEventListener('keydown', escHandler);
    } catch (error) {
        console.error('[LTOA-Report] Erreur Vue par Client:', error);
        alert('Erreur lors de la génération de la vue par client. Consultez la console F12.');
    }
},

        // ============================================
        // VUE CHRONOLOGIQUE
        // ============================================
        showChronoView() {
    try {
        const {
            emailsSent = [], emailsAffected = [], aircallCalls = [],
            tasksCompleted = [], logs = [], estimates = [],
            policies = [], claims = [], user = '', date = ''
        } = this.data;

        const existing = document.getElementById('ltoa-chrono-modal');
        if (existing) existing.remove();

        const parseTime = value => {
            const match = String(value || '').match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
            if (!match) return null;
            const hours = parseInt(match[1], 10);
            const minutes = parseInt(match[2], 10);
            const seconds = parseInt(match[3] || '0', 10);
            return hours * 3600 + minutes * 60 + seconds;
        };

        const extractTime = (...values) => {
            for (const value of values) {
                const match = String(value || '').match(/(\d{1,2}):(\d{2})/);
                if (match) return `${match[1].padStart(2, '0')}:${match[2]}`;
            }
            return '';
        };

        const formatDuration = seconds => {
            if (!Number.isFinite(seconds) || seconds < 0) return '';
            if (seconds < 60) return `${seconds}s`;
            if (seconds < 3600) {
                const minutes = Math.floor(seconds / 60);
                const rest = seconds % 60;
                return rest ? `${minutes}min ${rest}s` : `${minutes}min`;
            }
            const hours = Math.floor(seconds / 3600);
            const minutes = Math.floor((seconds % 3600) / 60);
            return minutes ? `${hours}h ${minutes}min` : `${hours}h`;
        };

        const actions = [];
        const push = action => {
            const time = action.time || '';
            const seconds = parseTime(time);
            if (!time || seconds === null) return;
            actions.push({ ...action, timeSeconds: seconds });
        };

        emailsSent.forEach(item => push({
            category: 'emails',
            label: 'Email envoyé',
            time: extractTime(item.time, item.date),
            title: item.subject || 'Sans objet',
            client: item.clientName || item.recipientName || item.toEmail || item.to || '',
            detail: `À ${item.toEmail || item.to || '—'}`,
            summary: item.body || ''
        }));

        emailsAffected.forEach(item => push({
            category: 'emails',
            label: 'Email reçu / affecté',
            time: extractTime(item.time, item.date),
            title: item.subject || 'Sans objet',
            client: item.clientName || item.fromName || item.from || item.fromEmail || '',
            detail: `De ${item.from || item.fromEmail || '—'}${item.affectedTo ? ` · affecté à ${item.affectedTo}` : ''}`,
            summary: item.body || ''
        }));

        aircallCalls.forEach(item => push({
            category: 'calls',
            label: item.type === 'sortant' ? 'Appel sortant' : 'Appel entrant',
            time: extractTime(item.time),
            title: item.contact || item.phone || 'Contact inconnu',
            client: item.clientName || item.contact || '',
            detail: `${item.duration || '0s'}${item.mood ? ` · ${item.mood}` : ''}`,
            summary: item.summary || '',
            extra: [
                item.topics?.length ? `Sujets : ${item.topics.map(topic => typeof topic === 'string' ? topic : (topic?.name || topic?.content || '')).filter(Boolean).join(', ')}` : '',
                item.actionItems?.length ? `Actions : ${item.actionItems.join(' · ')}` : ''
            ].filter(Boolean).join('\n')
        }));

        tasksCompleted.forEach(item => push({
            category: 'tasks',
            label: 'Tâche terminée',
            time: extractTime(item.closedTime, item.time, item.completedDate),
            title: item.title || 'Tâche',
            client: item.clientName || item.client || '',
            detail: item.closedBy ? `Clôturée par ${item.closedBy}` : 'Tâche clôturée',
            summary: item.content || ''
        }));

        estimates.forEach(item => {
            const classification = ActivityDictionary.classify(item);
            push({
                category: 'estimates',
                label: `Devis · ${classification.label}`,
                time: extractTime(item.date),
                title: `Devis n° ${item.entityId || '—'}`,
                client: item.clientName || item.entityName || '',
                detail: ActivityDictionary.summarize(item),
                summary: ''
            });
        });

        policies.forEach(item => {
            const classification = ActivityDictionary.classify(item);
            push({
                category: 'policies',
                label: `Contrat · ${classification.label}`,
                time: extractTime(item.date),
                title: `Contrat n° ${item.entityId || '—'}`,
                client: item.clientName || item.entityName || '',
                detail: ActivityDictionary.summarize(item),
                summary: ''
            });
        });

        claims.forEach(item => {
            const classification = ActivityDictionary.classify(item);
            push({
                category: 'claims',
                label: `Sinistre · ${classification.label}`,
                time: extractTime(item.date),
                title: `Sinistre n° ${item.entityId || '—'}`,
                client: item.clientName || item.entityName || '',
                detail: ActivityDictionary.summarize(item),
                summary: ''
            });
        });

        logs.forEach(item => push({
            category: 'other',
            label: item.table || item.tableRaw || 'Autre action',
            time: extractTime(item.date),
            title: item.entityName || item.action || item.actionRaw || 'Action',
            client: item.clientName || '',
            detail: ActivityDictionary.summarize(item),
            summary: ''
        }));

        actions.sort((a, b) => a.timeSeconds - b.timeSeconds);

        const counts = actions.reduce((acc, action) => {
            acc[action.category] = (acc[action.category] || 0) + 1;
            return acc;
        }, {});

        const categoryLabels = {
            emails: 'Emails',
            calls: 'Appels',
            tasks: 'Tâches',
            estimates: 'Devis',
            policies: 'Contrats',
            claims: 'Sinistres',
            other: 'Autres'
        };

        const first = actions[0] || null;
        const last = actions[actions.length - 1] || null;
        const spanSeconds = first && last ? Math.max(0, last.timeSeconds - first.timeSeconds) : 0;

        const actionRows = actions.map((action, index) => {
            const previous = index > 0 ? actions[index - 1] : null;
            const elapsed = previous ? action.timeSeconds - previous.timeSeconds : 0;
            const search = String(`${action.title} ${action.client} ${action.detail} ${action.label}`).toLowerCase();

            return `
                <article class="ltoa-time-row" data-category="${action.category}" data-search="${Utils.escapeHtml(search)}">
                    <div class="ltoa-time-col">
                        <strong>${Utils.escapeHtml(action.time)}</strong>
                        ${elapsed > 60 ? `<span>+${Utils.escapeHtml(formatDuration(elapsed))}</span>` : ''}
                    </div>
                    <div class="ltoa-time-line"><i></i></div>
                    <div class="ltoa-time-card">
                        <div class="ltoa-time-card-head">
                            <span class="ltoa-time-tag">${Utils.escapeHtml(action.label)}</span>
                            ${action.client ? `<span class="ltoa-time-client">${Utils.escapeHtml(action.client)}</span>` : ''}
                        </div>
                        <strong class="ltoa-time-title">${Utils.escapeHtml(action.title)}</strong>
                        ${action.detail ? `<div class="ltoa-time-detail">${Utils.escapeHtml(action.detail)}</div>` : ''}
                        ${action.summary || action.extra ? `
                            <details class="ltoa-time-more">
                                <summary>Voir le détail</summary>
                                ${action.summary ? `<div>${Utils.escapeHtml(Utils.cleanRichText(action.summary))}</div>` : ''}
                                ${action.extra ? `<div>${Utils.escapeHtml(action.extra)}</div>` : ''}
                            </details>` : ''}
                    </div>
                </article>`;
        }).join('');

        const filterButtons = Object.entries(categoryLabels)
            .filter(([key]) => counts[key])
            .map(([key, label]) => `<button type="button" class="ltoa-chrono-filter" data-filter="${key}">${label}<b>${counts[key]}</b></button>`)
            .join('');

        const html = `
            <div id="ltoa-chrono-modal">
                <style>
                    #ltoa-chrono-modal{position:fixed;inset:0;z-index:2147483647;background:#f5f7fb;color:#101828;font-family:Inter,-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;overflow:auto}
                    #ltoa-chrono-modal *{box-sizing:border-box}
                    .ltoa-chrono-top{position:sticky;top:0;z-index:10;background:rgba(248,250,253,.92);backdrop-filter:blur(16px);border-bottom:1px solid #e5e9ef;padding:13px 22px;display:flex;align-items:center;justify-content:space-between;gap:16px}.ltoa-chrono-title h2{margin:0;font-size:25px;letter-spacing:-.03em}.ltoa-chrono-title p{margin:4px 0 0;color:#7a8695;font-size:12px}.ltoa-chrono-close{border:1px solid #dfe4ea;background:#fff;border-radius:999px;padding:9px 12px;font-size:12px;color:#344054;cursor:pointer}
                    .ltoa-chrono-main{max-width:1050px;margin:0 auto;padding:22px 24px 44px}.ltoa-chrono-summary{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;margin-bottom:13px}.ltoa-chrono-stat{background:#fff;border:1px solid #e4e8ed;border-radius:17px;padding:14px 16px}.ltoa-chrono-stat strong{display:block;font-size:22px}.ltoa-chrono-stat span{font-size:10px;color:#7a8695}
                    .ltoa-chrono-tools{display:flex;gap:7px;align-items:center;flex-wrap:wrap;background:#fff;border:1px solid #e4e8ed;border-radius:17px;padding:10px;margin-bottom:16px}.ltoa-chrono-search{flex:1;min-width:220px;border:0;background:#f6f8fb;border-radius:11px;padding:10px 11px;font:inherit;font-size:11px;outline:none}.ltoa-chrono-filter{border:1px solid #e4e8ed;background:#fff;border-radius:999px;padding:7px 9px;font-size:10px;color:#526071;cursor:pointer}.ltoa-chrono-filter b{margin-left:5px}.ltoa-chrono-filter.active{background:#eef3ff;border-color:#d5e0ff;color:#315ea8}
                    .ltoa-timeline{position:relative}.ltoa-time-row{display:grid;grid-template-columns:72px 20px minmax(0,1fr);gap:8px;align-items:stretch;margin-bottom:8px}.ltoa-time-col{text-align:right;padding-top:11px}.ltoa-time-col strong{display:block;font-size:11px}.ltoa-time-col span{display:block;font-size:8px;color:#98a2b3;margin-top:3px}.ltoa-time-line{position:relative}.ltoa-time-line:before{content:"";position:absolute;left:9px;top:0;bottom:-9px;width:1px;background:#dfe5ec}.ltoa-time-line i{position:absolute;left:5px;top:14px;width:9px;height:9px;border-radius:50%;background:#5d7fd8;box-shadow:0 0 0 4px #edf3ff}
                    .ltoa-time-card{background:#fff;border:1px solid #e4e8ed;border-radius:15px;padding:11px 12px;box-shadow:0 7px 22px rgba(15,23,42,.025)}.ltoa-time-card-head{display:flex;justify-content:space-between;gap:9px;align-items:center;margin-bottom:5px}.ltoa-time-tag{font-size:9px;background:#f1f4f8;border-radius:999px;padding:4px 7px;color:#596579;font-weight:700}.ltoa-time-row[data-category="emails"] .ltoa-time-tag{background:#ecfeff;color:#0e7490}.ltoa-time-row[data-category="calls"] .ltoa-time-tag{background:#fff7ed;color:#c2410c}.ltoa-time-row[data-category="tasks"] .ltoa-time-tag{background:#fffbeb;color:#b45309}.ltoa-time-row[data-category="estimates"] .ltoa-time-tag{background:#eff6ff;color:#1d4ed8}.ltoa-time-row[data-category="policies"] .ltoa-time-tag{background:#f5f3ff;color:#7c3aed}.ltoa-time-row[data-category="claims"] .ltoa-time-tag{background:#fef2f2;color:#b91c1c}.ltoa-time-client{font-size:9px;color:#667085}.ltoa-time-title{display:block;font-size:12px}.ltoa-time-detail{font-size:10px;color:#667085;margin-top:4px}.ltoa-time-more{margin-top:8px;border-top:1px solid #edf0f3;padding-top:7px}.ltoa-time-more summary{cursor:pointer;color:#315ea8;font-size:9px;font-weight:700}.ltoa-time-more div{white-space:pre-wrap;font-size:10px;color:#4d5968;line-height:1.5;margin-top:7px;background:#f7f9fb;border-radius:9px;padding:8px}
                    .ltoa-chrono-empty{background:#fff;border:1px solid #e4e8ed;border-radius:18px;padding:42px;text-align:center;color:#7a8695;font-size:12px}
                    @media(max-width:760px){.ltoa-chrono-main{padding:14px}.ltoa-chrono-summary{grid-template-columns:1fr 1fr}.ltoa-time-row{grid-template-columns:48px 16px minmax(0,1fr)}.ltoa-time-col strong{font-size:10px}.ltoa-time-line:before{left:7px}.ltoa-time-line i{left:3px}.ltoa-time-card-head{align-items:flex-start;flex-direction:column}.ltoa-chrono-top{align-items:flex-start}}
                </style>
                <header class="ltoa-chrono-top">
                    <div class="ltoa-chrono-title">
                        <h2>Chronologie</h2>
                        <p>${Utils.escapeHtml(user)} · ${Utils.escapeHtml(date)} · ordre réel des actions horodatées</p>
                    </div>
                    <button type="button" id="ltoa-close-chrono" class="ltoa-chrono-close">Fermer</button>
                </header>
                <main class="ltoa-chrono-main">
                    <div class="ltoa-chrono-summary">
                        <div class="ltoa-chrono-stat"><strong>${actions.length}</strong><span>actions horodatées</span></div>
                        <div class="ltoa-chrono-stat"><strong>${first ? Utils.escapeHtml(first.time) : '—'}</strong><span>première action</span></div>
                        <div class="ltoa-chrono-stat"><strong>${last ? Utils.escapeHtml(last.time) : '—'}</strong><span>dernière action</span></div>
                        <div class="ltoa-chrono-stat"><strong>${actions.length > 1 ? Utils.escapeHtml(formatDuration(spanSeconds)) : '—'}</strong><span>plage entre première et dernière</span></div>
                    </div>
                    <div class="ltoa-chrono-tools">
                        <input id="ltoa-chrono-search" class="ltoa-chrono-search" type="search" placeholder="Rechercher une action ou un client…">
                        <button type="button" class="ltoa-chrono-filter active" data-filter="all">Tout<b>${actions.length}</b></button>
                        ${filterButtons}
                    </div>
                    <div id="ltoa-timeline" class="ltoa-timeline">
                        ${actionRows || '<div class="ltoa-chrono-empty">Aucune action horodatée disponible pour cette journée.</div>'}
                    </div>
                </main>
            </div>`;

        document.body.insertAdjacentHTML('beforeend', html);
        const modal = document.getElementById('ltoa-chrono-modal');
        const searchInput = document.getElementById('ltoa-chrono-search');
        let currentFilter = 'all';

        const applyFilters = () => {
            const query = String(searchInput?.value || '').toLowerCase().trim();
            modal.querySelectorAll('.ltoa-time-row').forEach(row => {
                const categoryOk = currentFilter === 'all' || row.dataset.category === currentFilter;
                const searchOk = !query || String(row.dataset.search || '').includes(query);
                row.style.display = categoryOk && searchOk ? '' : 'none';
            });
        };

        modal.querySelectorAll('.ltoa-chrono-filter').forEach(button => {
            button.addEventListener('click', () => {
                currentFilter = button.dataset.filter || 'all';
                modal.querySelectorAll('.ltoa-chrono-filter').forEach(item => item.classList.remove('active'));
                button.classList.add('active');
                applyFilters();
            });
        });

        searchInput?.addEventListener('input', applyFilters);

        const close = () => modal?.remove();
        document.getElementById('ltoa-close-chrono')?.addEventListener('click', close);

        const escHandler = event => {
            if (event.key === 'Escape' && document.getElementById('ltoa-chrono-modal')) {
                close();
                document.removeEventListener('keydown', escHandler);
            }
        };
        document.addEventListener('keydown', escHandler);
    } catch (error) {
        console.error('[LTOA-Report] Erreur Vue Chronologique:', error);
        alert('Erreur lors de la génération de la vue chronologique. Consultez la console F12.');
    }
},

        exportHTML() {
            try {
                const source = document.getElementById('ltoa-report-modal');
                if (!source) {
                    alert('Le rapport doit être ouvert avant de pouvoir être exporté.');
                    return;
                }

                // L'export reprend directement l'interface actuellement affichée.
                // Cela évite d'avoir un ancien template HTML différent du rapport à l'écran.
                const clone = source.cloneNode(true);

                // Supprimer les contrôles purement interactifs du rapport.
                clone.querySelector('#ltoa-detail-backdrop')?.remove();
                clone.querySelector('#ltoa-view-by-client')?.remove();
                clone.querySelector('#ltoa-view-chrono')?.remove();
                clone.querySelector('#ltoa-export-html')?.remove();
                clone.querySelector('#ltoa-close-report')?.remove();

                // Toutes les catégories sont dépliées dans le fichier exporté afin que
                // les informations restent visibles même sans JavaScript.
                clone.querySelectorAll('details').forEach(details => {
                    details.setAttribute('open', '');
                });

                // Nettoyer les attributs interactifs inutiles.
                clone.querySelectorAll('button').forEach(button => {
                    if (button.classList.contains('ltoa-inline-filter') ||
                        button.classList.contains('ltoa-metric-btn') ||
                        button.classList.contains('ltoa-subtype-btn')) {
                        button.setAttribute('disabled', '');
                    }
                });

                const exportOverrides = document.createElement('style');
                exportOverrides.textContent = `
                    html,body{margin:0;padding:0;background:#f5f7fb!important}
                    #ltoa-report-modal{
                        position:static!important;
                        inset:auto!important;
                        min-height:100vh!important;
                        overflow:visible!important;
                        background:#f5f7fb!important;
                    }
                    .ltoa-topbar{position:static!important}
                    .ltoa-main{max-width:1500px!important}
                    .ltoa-chevron{display:none!important}
                    button[disabled]{cursor:default!important;opacity:1!important}
                    @media print{
                        #ltoa-report-modal{background:#fff!important}
                        .ltoa-topbar{position:static!important}
                        .ltoa-section{break-inside:avoid}
                    }
                `;
                clone.prepend(exportOverrides);

                const htmlContent = `<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Rapport d'activité - ${Utils.escapeHtml(this.data.user || '')} - ${Utils.escapeHtml(this.data.date || '')}</title>
</head>
<body>
${clone.outerHTML}
</body>
</html>`;

                const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
                const url = URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = `Rapport_${String(this.data.user || 'Utilisateur').replace(/\s+/g, '_')}_${String(this.data.date || '').replace(/\//g, '-')}.html`;
                document.body.appendChild(link);
                link.click();
                link.remove();
                URL.revokeObjectURL(url);
            } catch (error) {
                console.error('[LTOA-Report] Erreur export HTML:', error);
                alert('Erreur lors de l’export HTML. Consultez la console F12.');
            }
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

            // Charge actuelle : toutes les tâches ouvertes affectées au collaborateur
            const pendingTasks = await PendingTasksCollector.collect(userId, connectedUser, loader.updateStatus);
            await Utils.delay(CONFIG.DELAY_BETWEEN_REQUESTS);

            // Charge actuelle en devis depuis le tableau de bord
            const assignedEstimates = await EstimatesAssignmentCollector.collect(userId, loader.updateStatus);
            await Utils.delay(CONFIG.DELAY_BETWEEN_REQUESTS);

            // Étape 3: Appels Aircall
            loader.update(3, 20, 'Collecte des appels Aircall...');
            let aircallCalls = [];
            if (CONFIG.AIRCALL_ENABLED) {
                try {
                    aircallCalls = await AircallCollector.collect(connectedUser, Utils.getTodayDate(), loader.updateStatus);
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
            // On dérive ce compteur de la liste exhaustive des tâches ouvertes afin
            // d'éviter deux collectes différentes qui donnaient des totaux incohérents.
            loader.update(5, 48, 'Classement des tâches en retard...');
            const tasksOverdue = pendingTasks
                .filter(task => task.isOverdue)
                .sort((a, b) => (b.daysOverdue || 0) - (a.daysOverdue || 0));

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
                assignedEstimates: assignedEstimates, // Devis actuellement assignés + répartition par état
                aircallCalls: aircallCalls, // Ajouter les appels Aircall
                tasksCompleted: resolvedData.tasksCompleted,
                tasksOverdue: resolvedData.tasksOverdue,
                pendingTasks: pendingTasks,
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
            const existing = document.getElementById('ltoa-date-picker-modal');
            if (existing) existing.remove();

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
                <style>
                    #ltoa-date-picker-modal *{box-sizing:border-box}
                    .ltoa-picker-backdrop{
                        position:fixed;inset:0;z-index:2147483646;
                        display:flex;align-items:center;justify-content:center;padding:24px;
                        background:rgba(15,23,42,.28);backdrop-filter:blur(8px);
                        font-family:Inter,-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;
                    }
                    .ltoa-picker-card{
                        width:min(620px,96vw);
                        background:linear-gradient(180deg,rgba(255,255,255,.98),rgba(250,252,255,.96));
                        border:1px solid rgba(255,255,255,.86);
                        border-radius:26px;padding:24px;
                        box-shadow:0 30px 90px rgba(15,23,42,.20);
                        color:#101828;
                    }
                    .ltoa-picker-head{display:flex;align-items:flex-start;justify-content:space-between;gap:16px;margin-bottom:22px}
                    .ltoa-picker-kicker{font-size:10px;text-transform:uppercase;letter-spacing:.09em;color:#98a2b3;margin-bottom:5px}
                    .ltoa-picker-head h3{margin:0;font-size:24px;letter-spacing:-.03em;font-weight:760}
                    .ltoa-picker-head p{margin:5px 0 0;font-size:12px;color:#748091}
                    .ltoa-picker-close{border:0;background:#f1f4f7;width:36px;height:36px;border-radius:999px;cursor:pointer;color:#667085;font-size:18px}
                    .ltoa-picker-label{display:block;font-size:11px;font-weight:700;color:#475467;margin:0 0 8px}
                    .ltoa-picker-input,.ltoa-picker-note{
                        width:100%;border:1px solid #dfe4ea;background:#fff;border-radius:14px;
                        padding:12px 13px;font:inherit;color:#101828;outline:none;
                        transition:border-color .16s,box-shadow .16s;
                    }
                    .ltoa-picker-input:focus,.ltoa-picker-note:focus{
                        border-color:#9ab8ff;box-shadow:0 0 0 4px rgba(37,99,235,.08)
                    }
                    .ltoa-picker-note{min-height:108px;resize:vertical;line-height:1.5}
                    .ltoa-picker-field{margin-bottom:18px}
                    .ltoa-quick-dates{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-top:9px}
                    .ltoa-quick-date{
                        border:1px solid #e3e8ef;background:#fff;border-radius:13px;padding:10px 12px;
                        cursor:pointer;color:#475467;text-align:left;font:inherit;transition:.16s ease
                    }
                    .ltoa-quick-date:hover{transform:translateY(-1px);border-color:#cbd8f5;box-shadow:0 8px 20px rgba(37,99,235,.06)}
                    .ltoa-quick-date.active{background:#f3f7ff;border-color:#cddcff;color:#1d4ed8}
                    .ltoa-quick-date strong{display:block;font-size:12px;margin-bottom:3px}
                    .ltoa-quick-date small{font-size:10px;color:#8b96a5}
                    .ltoa-picker-hint{font-size:10px;color:#98a2b3;margin-top:7px}
                    .ltoa-picker-actions{display:flex;justify-content:flex-end;gap:9px;margin-top:22px}
                    .ltoa-picker-btn{border:1px solid #dfe4ea;background:#fff;color:#475467;border-radius:999px;padding:10px 15px;font:inherit;font-size:12px;font-weight:650;cursor:pointer}
                    .ltoa-picker-btn.primary{border-color:transparent;background:linear-gradient(180deg,#2c6bff,#2259da);color:#fff;box-shadow:0 14px 30px rgba(37,99,235,.22)}
                    @media(max-width:620px){.ltoa-quick-dates{grid-template-columns:1fr}.ltoa-picker-card{padding:18px}.ltoa-picker-actions{flex-direction:column-reverse}.ltoa-picker-btn{width:100%}}
                </style>
                <div class="ltoa-picker-backdrop">
                    <div class="ltoa-picker-card" role="dialog" aria-modal="true" aria-labelledby="ltoa-picker-title">
                        <div class="ltoa-picker-head">
                            <div>
                                <div class="ltoa-picker-kicker">Rapport d’activité</div>
                                <h3 id="ltoa-picker-title">Générer le rapport</h3>
                                <p>Choisissez la journée à analyser et ajoutez une note si nécessaire.</p>
                            </div>
                            <button type="button" class="ltoa-picker-close" id="ltoa-picker-close" aria-label="Fermer">×</button>
                        </div>

                        <div class="ltoa-picker-field">
                            <label class="ltoa-picker-label" for="ltoa-date-input">Date du rapport</label>
                            <input class="ltoa-picker-input" type="date" id="ltoa-date-input" value="${formatDateInput(today)}" max="${formatDateInput(today)}">
                            <div class="ltoa-quick-dates">
                                <button type="button" class="ltoa-quick-date active" data-date="${formatDate(today)}"><strong>Aujourd’hui</strong><small>${formatDate(today)}</small></button>
                                <button type="button" class="ltoa-quick-date" data-date="${formatDate(yesterday)}"><strong>Hier</strong><small>${formatDate(yesterday)}</small></button>
                                <button type="button" class="ltoa-quick-date" data-date="${formatDate(twoDaysAgo)}"><strong>Avant-hier</strong><small>${formatDate(twoDaysAgo)}</small></button>
                            </div>
                        </div>

                        <div class="ltoa-picker-field">
                            <label class="ltoa-picker-label" for="ltoa-report-notes">Note ou précision complémentaire <span style="font-weight:500;color:#98a2b3">· facultatif</span></label>
                            <textarea class="ltoa-picker-note" id="ltoa-report-notes" placeholder="Ex. rendez-vous extérieur, travail de fond, incident technique, précision sur un dossier…">${Utils.escapeHtml(REPORT_NOTES)}</textarea>
                            <div class="ltoa-picker-hint">La note apparaîtra en haut du rapport uniquement si elle est renseignée.</div>
                        </div>

                        <div class="ltoa-picker-actions">
                            <button type="button" id="ltoa-date-cancel" class="ltoa-picker-btn">Annuler</button>
                            <button type="button" id="ltoa-date-confirm" class="ltoa-picker-btn primary">Générer le rapport</button>
                        </div>
                    </div>
                </div>
            `;

            document.body.appendChild(modal);

            const dateInput = document.getElementById('ltoa-date-input');
            const noteInput = document.getElementById('ltoa-report-notes');
            dateInput.focus();

            const cancel = () => {
                modal.remove();
                resolve(null);
            };

            modal.querySelectorAll('.ltoa-quick-date').forEach(btn => {
                btn.addEventListener('click', () => {
                    const parts = btn.dataset.date.split('/');
                    dateInput.value = `${parts[2]}-${parts[1]}-${parts[0]}`;
                    modal.querySelectorAll('.ltoa-quick-date').forEach(item => item.classList.remove('active'));
                    btn.classList.add('active');
                });
            });

            dateInput.addEventListener('change', () => {
                modal.querySelectorAll('.ltoa-quick-date').forEach(item => {
                    const parts = item.dataset.date.split('/');
                    const inputValue = `${parts[2]}-${parts[1]}-${parts[0]}`;
                    item.classList.toggle('active', inputValue === dateInput.value);
                });
            });

            document.getElementById('ltoa-date-cancel').addEventListener('click', cancel);
            document.getElementById('ltoa-picker-close').addEventListener('click', cancel);

            document.getElementById('ltoa-date-confirm').addEventListener('click', () => {
                const inputValue = dateInput.value;
                if (!inputValue) return;
                REPORT_NOTES = noteInput?.value.trim() || '';
                const parts = inputValue.split('-');
                const formattedDate = `${parts[2]}/${parts[1]}/${parts[0]}`;
                modal.remove();
                resolve(formattedDate);
            });

            modal.querySelector('.ltoa-picker-backdrop').addEventListener('click', (e) => {
                if (e.target === e.currentTarget) cancel();
            });

            modal.addEventListener('keydown', (e) => {
                if (e.key === 'Escape') cancel();
                if (e.key === 'Enter' && e.target !== noteInput) {
                    document.getElementById('ltoa-date-confirm').click();
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
