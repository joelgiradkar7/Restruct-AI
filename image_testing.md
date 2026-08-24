# Image integration testing playbook
- Use only base64-encoded JPEG, PNG, or WEBP images.
- Test with a non-blank image containing real visual features.
- Reject SVG, BMP, HEIC, GIF/APNG, blank, or uniform images.
- Re-detect MIME type after any image transformation.
- Keep image payloads reasonably sized.