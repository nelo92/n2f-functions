# n2f-functions

Creator MAU

## Firebase url

https://console.firebase.google.com/project/nodedefrais/overview

## API

1. Environnement production

https://europe-west1-nodedefrais.cloudfunctions.net/api

[Call isUp](https://europe-west1-nodedefrais.cloudfunctions.net/api/isup)

2. Environnement Local

   http://localhost:5001/nodedefrais/europe-west1/api

   [Call isUp](http://localhost:5001/nodedefrais/europe-west1/api/isUp)

## Project configuration

Need file <nodedefrais-firebase-adminsdk.json> and put it in directory functions/firebase

## Installation et exécution

Avant de lancer les scripts, il faut se placer dans le répertoire functions puis installer les dépendances :

```bash
cd functions
npm install
```

Scripts disponibles :

- `npm run lint` : vérifie le code avec ESLint.
- `npm run serve` : démarre les émulateurs Firebase pour tester les Cloud Functions localement.
- `npm run shell` : ouvre un shell Firebase pour tester les fonctions.
- `npm run start` : lance le shell Firebase.
- `npm run deploy` : déploie les fonctions sur Firebase.
- `npm run logs` : affiche les logs des fonctions.

## Version

## Firebase reference

Create and Deploy Your First Cloud Functions
https://firebase.google.com/docs/functions/write-firebase-functions
