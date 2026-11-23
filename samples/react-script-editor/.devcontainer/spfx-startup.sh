set -e

echo
echo -e "\e[1;94mInstalling Node dependencies\e[0m"
npm install

## commands to create dev certificate and copy it to the root folder of the project
echo
echo -e "\e[1;94mGenerating dev certificate\e[0m"
npx heft trust-dev-cert

# Export the development CA certificate for trust on the host.
openssl x509 -inform PEM -in ~/.rushstack/rushstack-ca.pem -outform DER -out ./spfx-dev-cert.cer
cp ~/.rushstack/rushstack-ca.pem ./spfx-dev-cert.pem

echo
echo -e "\e[1;92mReady!\e[0m"

echo -e "\n\e[1;94m**********\nOptional: if you plan on using heft start, don't forget to add the container certificate to your local machine. Please visit https://aka.ms/spfx-devcontainer for more information\n**********"