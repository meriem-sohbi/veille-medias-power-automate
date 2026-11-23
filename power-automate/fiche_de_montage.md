# Veille médias hebdomadaire, fiche de montage Power Automate

Cette fiche correspond au rapport « Synthèse hebdomadaire des tendances médias ». Elle permet de monter le flux en 16 actions, en 45 minutes environ, dans un environnement Microsoft 365. La logique (période, doublons, contrôles, mise en forme) est dans trois Office Scripts déjà écrits et testés localement. Le flux ne fait que les enchaîner.

Les noms des actions sont donnés en français, avec le nom anglais entre parenthèses.

## Contenu du dépôt

| Fichier | Rôle |
|---|---|
| excel/Synthese_hebdo_medias.xlsx | Classeur partagé. Tables Parametres, Sources (liste blanche), Audiences (export interne), Brouillon, Contexte (masquée), Suivi. Données fictives. |
| office-scripts/1_Preparer.ts | Garde les sept derniers jours, retire les mots exclus, repère les doublons par l’URL nettoyée ou le titre, signale les chiffres qui diffèrent d’une source à l’autre. Prépare le texte envoyé à l’IA : titre, extrait, date et source, sans adresse web ni donnée interne. |
| office-scripts/2_Controler.ts | Lit la réponse de l’IA (catégorie, résumé, points à vérifier). Vérifie que chaque chiffre et chaque nom propre existe dans la source, écarte les textes contenant une consigne cachée, marque « à arbitrer » ou « à confirmer », calcule les audiences sans IA et écrit le brouillon dans la table Brouillon. |
| office-scripts/3_Publier.ts | Après approbation, lit le brouillon corrigé, produit la version validée avec le nom de la personne et ajoute une ligne au suivi du pilote. Refuse toute publication sans nom. |
| power-automate/prompt_ai_builder.txt | Texte du prompt, avec la consigne du rapport. |
| demo/ | Démonstration dans le navigateur, qui exécute les trois scripts sur les données fictives. |
| tests/ | Banc d’essai local, cinq scénarios (« npm test »). |

## Prérequis à vérifier avec l’administrateur Power Platform

| Brique | Connecteur | Niveau |
|---|---|---|
| Lecture des flux | RSS | Standard |
| Scripts et classeur | Excel Online (Business), action Exécuter un script | Standard. Office Scripts doit être activé dans le tenant. |
| Validation | Approbations (Approvals) | Standard |
| Diffusion et alertes | Microsoft Teams | Standard |
| Résumés IA | AI Builder, action Exécuter un prompt | Consomme des crédits AI Builder ou Copilot. Licence à confirmer. |
| Connexion directe à l’API interne | HTTP ou connecteur personnalisé | Premium. Solution de repli de la V1, un export régulier des audiences dans la table Audiences. |

Les règles DLP du tenant doivent autoriser ces connecteurs ensemble.

## Préparation

1. Déposer excel/Synthese_hebdo_medias.xlsx dans la bibliothèque SharePoint de l’équipe.
2. L’ouvrir dans Excel pour le web, onglet Automatiser, Nouveau script. Coller office-scripts/1_Preparer.ts et enregistrer sous le nom 1_Preparer. Faire de même pour 2_Controler et 3_Publier.
3. Remplir la table Sources avec les vrais flux et la table Parametres avec les vrais mots-clés.
4. Créer le prompt « Synthese_hebdo_veille » à partir de power-automate/prompt_ai_builder.txt, et le tester avec power-automate/exemple_elements.json.

## Les 16 actions du flux

Flux de cloud planifié, nommé « Veille médias hebdomadaire ». Les actions 3 à 15 sont regroupées dans une étendue nommée Principal.

| N° | Action | Réglages |
|---|---|---|
| 1 | Périodicité (Recurrence) | Chaque semaine, lundi, 6 h 30, fuseau Paris. |
| 2 | Initialiser une variable (deux fois) | articles, type Tableau, vide. echecs, type Tableau, vide. |
| 3 | Excel, Répertorier les lignes présentes dans une table (List rows present in a table) | Classeur partagé, table Sources. Filtre : Active eq 'Oui'. |
| 4 | Appliquer à chacun (Apply to each), nommé Pour_chaque_source | Sur la valeur de l’action 3. Contient les actions 5 à 8. |
| 5 | RSS, Répertorier tous les éléments de flux RSS (List all RSS feed items) | URL = colonne URL du flux. Paramètres, stratégie de nouvelle tentative : intervalle fixe, 3 essais, PT1M. |
| 6 | Sélectionner (Select) | Depuis la sortie RSS. Mappage ci-dessous. |
| 7 | Ajouter à la variable de tableau (Append to array variable), articles | Valeur = sortie de Sélectionner. |
| 8 | Ajouter à la variable de tableau, echecs | Configurer l’exécution après : l’action 5 a échoué ou a expiré. Valeur ci-dessous. |
| 9 | Excel, Exécuter un script (Run script), 1_Preparer | articlesJson = string(variables('articles')). dateReference = utcNow(). |
| 10 | AI Builder, Exécuter un prompt (Run a prompt), dans une étendue nommée IA | Prompt Synthese_hebdo_veille. elements = json(outputs('1_Preparer')?['body/result'])?['elementsIA']. |
| 11 | Excel, Exécuter un script, 2_Controler | Exécuter après l’étendue IA, qu’elle ait réussi, échoué ou expiré. preparationJson = résultat de l’action 9. reponseIA = texte renvoyé par le prompt. sourcesEnEchecJson = string(variables('echecs')). |
| 12 | Condition, alerte de collecte | Si nbSourcesEnEchec (sortie de l’action 11) est supérieur à 0 : Teams, Publier un message dans une conversation ou un canal (Post message in a chat or channel), « Sources non collectées cette semaine » avec la liste des échecs. |
| 13 | Approbations, Démarrer et attendre une approbation (Start and wait for an approval) | Type Approuver/Refuser, premier à répondre. Affecté au responsable éditorial. Détails = resumeApprobation (sortie de l’action 11). Lien de l’élément = lien du classeur. |
| 14 | Condition, réponse | Résultat de l’approbation égal à Approve. Si oui : Excel, Exécuter un script, 3_Publier, avec les paramètres ci-dessous. |
| 15 | Teams, Publier un message dans une conversation ou un canal | Si oui : message = html (sortie de 3_Publier) dans le canal de l’équipe. Si non : message à la personne qui prépare, avec le commentaire du refus. La préparation manuelle reste possible. |
| 16 | Étendue Alerte, exécutée si Principal a échoué ou expiré | Teams, message à l’équipe avec le lien de l’exécution (expression ci-dessous). |

Mappage de l’action 6, Sélectionner.

```
titre     item()?['title']
url       item()?['primaryLink']
extrait   item()?['summary']
date      item()?['publishDate']
source    items('Pour_chaque_source')?['Nom']
fiabilite items('Pour_chaque_source')?['Fiabilite']
```

Valeur de l’action 8.

```
{ "source": "@{items('Pour_chaque_source')?['Nom']}", "erreur": "indisponible après 3 tentatives" }
```

Paramètres de 3_Publier, action 14.

```
validePar          first(outputs('Démarrer_et_attendre_une_approbation')?['body/responses'])?['responder/displayName']
commentaire        first(outputs('Démarrer_et_attendre_une_approbation')?['body/responses'])?['comments']
dureeRelectureMin  durée entre la demande d’approbation et la réponse, en minutes. Choisir les champs
                   « Date de la demande » et « Date de réponse » dans le contenu dynamique, puis :
                   div(sub(ticks(<date de réponse>), ticks(<date de la demande>)), 600000000)
```

Lien de l’exécution, action 16.

```
concat('https://make.powerautomate.com/environments/', workflow()?['tags']?['environmentName'], '/flows/', workflow()?['name'], '/runs/', workflow()?['run']?['name'])
```

Les noms internes des actions servent dans les expressions. Si une action porte un autre nom, il faut adapter l’expression ou choisir le champ dans le contenu dynamique.

## Relecture par le responsable

La demande d’approbation arrive dans Teams ou Outlook avec un lien vers le classeur. Dans la table Brouillon, le responsable corrige la colonne Resume, met Non dans Garder pour retirer un élément, puis approuve. Les colonnes A verifier, Controles et Statut indiquent ce qui doit être regardé en priorité.

## Tester

1. Remplacer les adresses fictives par un ou deux vrais flux et lancer le flux à la main.
2. Vérifier la table Brouillon, puis approuver depuis Teams ou Outlook.
3. Vérifier le message dans le canal et la nouvelle ligne de la table Suivi.
4. Ajouter une adresse de flux erronée pour vérifier l’alerte de collecte, et vider la table Audiences pour vérifier l’affichage « indisponibles ».

## Ce qui a été testé et ce qui ne l’a pas été

Testé localement, avec un faux classeur qui reproduit les méthodes Excel utilisées et les données fictives du rapport. Les trois scripts compilent en TypeScript strict, sans type any, comme l’exige Office Scripts. Cinq scénarios passent : nominal avec relecture, réponse IA illisible, IA qui invente un chiffre et un nom, audiences absentes, réponse IA entourée de balises de code. Une validation sans nom est refusée. Aucune adresse web ni donnée interne ne part vers l’IA.

Non testé, faute d’accès à un environnement Microsoft 365 : l’exécution réelle dans Excel et Power Automate, le prompt avec un vrai modèle, les vrais flux RSS, les noms exacts des champs dynamiques et les licences.
