# Canva Export + Magic Resize (Netlify)

Deploy on Netlify. Handles:
- OAuth with Canva (Connect API)
- Export original design to PPTX
- Magic-Resize (server-side) to 4:3, then export PPTX
- Optional S3 upload (if AWS env vars configured)

## Environment Variables (Netlify > Site settings > Environment)
- CANVA_CLIENT_ID=...
- CANVA_CLIENT_SECRET=...
- CANVA_REDIRECT=https://<your-site>.netlify.app/callback/canva
- SESSION_SECRET=<random-long-string>
- (optional) S3_BUCKET=...
- (optional) AWS_REGION=ap-southeast-1
- (optional) AWS_ACCESS_KEY_ID=...
- (optional) AWS_SECRET_ACCESS_KEY=...

## Run locally
netlify dev