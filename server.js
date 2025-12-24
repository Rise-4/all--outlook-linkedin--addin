const express = require('express');
const path = require('path');
const https = require('https');
const fs = require('fs');

const app = express();
const PORT = 3000;

// Servir les fichiers statiques
app.use(express.static(__dirname));

// Route pour la page d'accueil (si nécessaire)
app.get('/', (req, res) => {
  res.send('Outlook Add-in Server is running');
});

// Options HTTPS (nécessaire pour les add-ins Office)
// Note: Pour le développement, vous devrez créer un certificat auto-signé
const httpsOptions = {
  key: fs.existsSync('localhost-key.pem') ? fs.readFileSync('localhost-key.pem') : null,
  cert: fs.existsSync('localhost.pem') ? fs.readFileSync('localhost.pem') : null
};

// Démarrer le serveur HTTPS si les certificats existent, sinon HTTP
if (httpsOptions.key && httpsOptions.cert) {
  https.createServer(httpsOptions, app).listen(PORT, () => {
    console.log(`Server running at https://localhost:${PORT}`);
    console.log('Add-in ready for testing!');
  });
} else {
  console.warn('⚠️  Certificats HTTPS non trouvés. Le serveur démarre en HTTP.');
  console.warn('⚠️  Pour les add-ins Office, HTTPS est requis. Créez des certificats avec:');
  console.warn('   openssl req -x509 -newkey rsa:2048 -keyout localhost-key.pem -out localhost.pem -days 365 -nodes');
  app.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
    console.log('⚠️  Note: Changez les URLs dans manifest.xml pour utiliser http:// au lieu de https://');
  });
}

