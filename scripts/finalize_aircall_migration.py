from pathlib import Path
import subprocess
import sys

helper = Path('scripts/migrate_aircall_netlify.py')
source = helper.read_text(encoding='utf-8')
lines = source.splitlines()
for i, line in enumerate(lines):
    if line.startswith('forbidden = '):
        lines[i] = "forbidden = ['api.aircall.io']"
        break
else:
    raise SystemExit('Migration guard not found')
helper.write_text('\n'.join(lines) + '\n', encoding='utf-8')

subprocess.run([sys.executable, str(helper)], check=True)

path = Path('LTOA-Modulr-Rapport-Quotidien.user.js')
text = path.read_text(encoding='utf-8')

# Replace only the old Aircall credential/configuration card.
anchor = text.index('id="ltoa-aircall-config-status"')
start = text.rfind('                        <div style="margin-bottom:20px;', 0, anchor)
end = text.index('                        <div style="display: flex; gap: 10px;', anchor)
static_card = '''                        <div style="margin-bottom:20px;padding:12px 14px;background:#f5f7fa;border:1px solid #e0e0e0;border-radius:8px;">
                            <div style="font-size:13px;font-weight:600;color:#333;">Connexion Aircall</div>
                            <div style="font-size:11px;color:#777;margin-top:3px;">Sécurisée via Netlify · aucune configuration nécessaire</div>
                        </div>

'''
text = text[:start] + static_card + text[end:]

# Remove only the old Aircall configure-button listener.
anchor = text.index("document.getElementById('ltoa-configure-aircall').addEventListener")
start = text.rfind('            ', 0, anchor)
end = text.index('            // Confirmer', anchor)
text = text[:start] + text[end:]

forbidden = [
    'ltoa_aircall_api_id',
    'ltoa_aircall_api_token',
    'api.aircall.io',
    'configureAircallApi',
    'Authorization: `Basic',
]
left = [value for value in forbidden if value in text]
if left:
    raise SystemExit(f'Aircall browser-secret references remain: {left}')
if 'https://aircallmodulr.netlify.app' not in text:
    raise SystemExit('Netlify endpoint missing')
if '// @version      5.3.0' not in text:
    raise SystemExit('Expected version 5.3.0 missing')

path.write_text(text, encoding='utf-8')
print('Aircall access migration finalized')
