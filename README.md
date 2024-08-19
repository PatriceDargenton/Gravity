# Gravity
Gravity Screen Saver : Un écran de veille chaotique
---

Gravity Screen Saver est un économiseur d'écran (un écran de veille pour Windows) sur l'aspect chaotique de la loi de gravité qui régie le mouvement des planètes. Il s'agit d'une simulation du mouvement harmonique (mais pas très stable !) de grappes de planètes parcourant des orbites purement théoriques. Ces planètes sont placées dans des conditions d'équilibre calculées à partir des racines unitaires complexes (Z<sup>n</sup> = 1). Il peut y avoir 2 systèmes imbriqués ! (par exemple une lune gravitant autour d'une terre, le tout gravitant autour d'un soleil). Ces orbites seraient stables à terme si la simulation numérique n'était pas une approximation des phénomènes continus de l'espace-temps, surtout lorsque deux planètes sont proches l'une de l'autre. N'hésitez pas à tester plusieurs fois pour tomber sur une figure intéressante, car toutes les figures sont aléatoires, mais paramétrables.

## Table des matières
- [Utilisation](#utilisation)
- [Limitations](#limitations)
- [Projets](#projets)
- [Versions](#versions)
- [Liens](#liens)

## Utilisation
- Les fichiers Gravity2.exe et Gravity2.exe.config étant renommés en Gravity2.scr et Gravity2.scr.config (c'est fait dans les évènements de post-build dans les options de compilation), il suffit de faire un double-clic sur le Gravity2.scr pour lancer l'écran de veille, et bouton droit : installer pour en faire un écran de veille reconnu par Windows ;
- Les images doivent être nommées Images\star*.bmp ou .jpg (les zones transparentes doivent être en vert, par exemple l'intérieur des anneaux de saturne) ;
- Les images de fond doivent être nommées Images\Space\*.jpg (elles sont facultatives, le nom n'a plus besoin de commencer par space, c'est plus simple d'utiliser un sous-dossier Space).

## Limitations
- La gestion des chocs ne fonctionne pas, du coup elle est désactivée et les corps célestes sont des fantômes les uns par rapport au autres, pour des distances négatives.

## Projets
- Version html en Blazor.

## Versions

Voir le [Changelog.md](Changelog.md)

## Liens

Documentation d'origine complète : [Gravity : index.html](http://patrice.dargenton.free.fr/gravity/index.html)