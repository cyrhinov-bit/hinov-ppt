# Présentation : HINOV GROUP — L'innovation au service de votre performance

Présentation institutionnelle et commerciale du **Groupe HINOV** — solutions en **technologies de l'information**, **infrastructures technologiques** et **communication**, basé à **Abidjan, Côte d'Ivoire**.

🌐 **Lien en ligne (Démo Web interactive)** : [https://cyrhinov-bit.github.io/datacenter/](https://cyrhinov-bit.github.io/datacenter/)

---

## 📋 Contenu du Projet

1. **`index.html`** :
   - Application web interactive de présentation (HTML5 / Tailwind CSS / Lucide Icons / Canvas Particles).
   - Contenu : présentation corporative **HINOV GROUP** (10 diapositives).
   - Modes de navigation :
     - Navigation manuelle (flèches clavier, boutons de navigation).
     - Défilement automatique (*Lecture Auto*) avec sélecteur de vitesse (**1.5s**, **2.5s**, **4.0s**, **6.0s**, **8.0s**, **10.0s**).
     - Mode plein écran (<kbd>F</kbd>).
     - Révélation progressive étape par étape des éléments.
     - Navigation tactile (gestes de balayage).

2. **Générateurs PowerPoint (`generate_pptx_style.py`)** :
   - Script Python (`python-pptx`) générant la présentation 16:9 au format `.pptx`.
   - Respect strict de la charte visuelle moderne (capsules, badges ronds, palette verte HINOV).
   - Logo : **`assets/logo.png`** (logo de l'entreprise HINOV GROUP).

3. **`assets/`** :
   - Logo : **`assets/logo.png`** (logo HINOV GROUP — vert `#4AD07B`).
   - Images issues du document source **`hinov.docx`** (une image par diapo, reflétant le contenu) :
     - `img_hinov_main.jpeg`, `img_vision.jpeg`, `img_mission.jpeg`, `img_expertise.jpeg`, `img_engagement.jpeg`.
   - Logos des partenaires (récupérés depuis `hinov.docx`) :
     - `partner_coris_bank.png`, `partner_medlog.png`, `partner_cit.jpeg`, `partner_pnoev.jpeg`, `partner_cnptir.jpeg`, `partner_mission_etrang.jpeg`, `partner_ecole_horizon.jpeg`, `partner_mahi.jpeg`, `partner_smart_security.jpeg`, `partner_tech_*.jpeg/png`.

---

## 🗂️ Structure des Diapositives (HINOV GROUP)

1. **Présentation Générale** — *L'innovation au service de votre performance.* (image : `assets/img_hinov_main.jpeg`)
2. **Notre Vision** — (image : `assets/img_vision.jpeg`)
3. **Notre Mission** — (image : `assets/img_mission.jpeg`)
4. **Domaines d'Expertise** — (image : `assets/img_expertise.jpeg`)
5. **Nos Engagements** — (image : `assets/img_engagement.jpeg`)
6. **Clients et Partenaires** — grille de logos des partenaires (récupérés depuis `hinov.docx`) : CORIS BANK, MEDLOG, CICT, PNOEV, CNPTIR, Ministère des Affaires Étrangères, École Académie Horizon International, Groupe Scolaire Mehi, Smart Security.
7. **Notre Approche de Travail** — (image : `assets/img_mission.jpeg`)
8. **Perspectives de Développement** — (image : `assets/img_vision.jpeg`)
9. **Pourquoi Choisir HINOV GROUP ?** — (image : `assets/img_expertise.jpeg`)
10. **Nos Coordonnées** — (image : `assets/img_hinov_main.jpeg`)

---

## 🚀 Utilisation

### 1. Visualiser la présentation Web interactive
Ouvrez simplement le fichier `index.html` dans n'importe quel navigateur moderne (Chrome, Edge, Firefox, Safari).

### 2. Régénérer le fichier PowerPoint (.pptx)
Installez les dépendances requises :
```bash
pip install python-pptx pillow lxml
```

Générez la présentation :
```bash
python generate_pptx_style.py
```
Le fichier `Presentation_Datacenter_Sante_CI_Style.pptx` sera généré à la racine.

---

## 🎨 Charte Graphique & Logo

- **Couleur principale** : Vert HINOV **`#4AD07B`** (issue du logo `assets/logo.png`).
- **Couleur secondaire / texte** : Bleu très foncé **`#102A38`**.
- **Arrière-plan** : Blanc / Gris très clair **`#F0F3F6`**.
- **Typographie** : Inter / Outfit.
- **Éléments de style** : Pills capsules, badges ronds, cartes arrondies, animations progressives.

---

## 🏛️ Crédits & Contexte

- **HINOV GROUP** — Abidjan, Côte d'Ivoire
- **Contact** : hinovgroup@hinovgroup.com | www.hinovgroup.com
- **Présentateur** : M. OUATTARA Yacouba
