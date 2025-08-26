FROM mcr.microsoft.com/powershell:7.4-alpine

WORKDIR /app

# Install dependencies
RUN apk add --no-cache git curl python3 py3-pip \
    && pip install gsutil

# Install Microsoft Graph PowerShell SDK
RUN pwsh -Command "Install-Module Microsoft.Graph -Force -Scope AllUsers -AllowClobber"

# Copy script
COPY MoveEncryptedEmail.ps1 .

CMD ["pwsh", "-File", "MoveEncryptedEmail.ps1"]
