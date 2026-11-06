// =========================================================
// LPT — le « cerveau » de l'appli.
// Plan du fichier :
//   1. Réglages de base et données
//   2. Stockage (Firebase, ou mode démo local)
//   3. Astuce du jour (le tirage)
//   4. Navigation entre les écrans
//   5. Affichage (accueil, fiche, formulaire, réglages)
//   6. Actions (favori, ajout, suppression...)
//   7. Export / import Excel
//   8. Démarrage
// =========================================================

import { initializeApp } from "https://www.gstatic.com/firebasejs/12.4.0/firebase-app.js";
import {
  getAuth, GoogleAuthProvider, signInWithPopup, signInWithRedirect,
  onAuthStateChanged, signOut
} from "https://www.gstatic.com/firebasejs/12.4.0/firebase-auth.js";
import {
  initializeFirestore, persistentLocalCache, persistentMultipleTabManager,
  doc, collection, onSnapshot, setDoc, writeBatch, serverTimestamp
} from "https://www.gstatic.com/firebasejs/12.4.0/firebase-firestore.js";
import { firebaseConfig } from "./firebase-config.js";
import { chargerExcelJS, ajouterFeuille, telechargerClasseur, lireClasseur, lireFeuille } from "./excel-style.js";

// ---------- 1. Réglages de base et données ----------

const VERSION = "v0.1.0";
const NOM_APPLI = "app-lpt";   // nom du dossier dans Firestore et des fichiers Excel

// Une astuce = { id, titre, texte, cat, lien, image, photo, favori, connue, vue, cree }
//   cat    : identifiant de sa catégorie (ex. "cuisine")
//   lien   : adresse de la page / vidéo d'où vient l'astuce (facultatif)
//   image  : adresse d'une image trouvée sur le web (facultatif)
//   photo  : photo prise avec le tel, réduite et rangée directement dans la fiche (facultatif)
//   favori : étoile ; connue : « Je la connais » = ne sort plus au tirage du jour
//   vue    : déjà tirée pendant le « tour » en cours (voir partie 3)
//   cree   : date d'ajout (en millisecondes), sert à trier la liste
let astuces = new Map();   // id -> astuce

// Données générales : { categories: [{ id, nom, emoji }], jour: { date, id } }
//   jour = l'astuce du jour, partagée entre le tel et le PC
const CATEGORIES_DEPART = [
  { id: "cuisine", nom: "Cuisine", emoji: "🍳" },
  { id: "mode", nom: "Mode", emoji: "👕" },
  { id: "quotidien", nom: "Quotidien", emoji: "🏠" },
  { id: "informatique", nom: "Informatique", emoji: "💻" },
  { id: "sante", nom: "Santé", emoji: "🌿" },
  { id: "argent", nom: "Argent", emoji: "💰" }
];
let donnees = { categories: structuredClone(CATEGORIES_DEPART), jour: null };

// Réglages de ressenti, gardés sur CET appareil seulement (pas synchronisés)
const CLE_REGLAGES = "lpt-reglages";
let reglages = { theme: "auto", animations: "completes" };
try { Object.assign(reglages, JSON.parse(localStorage.getItem(CLE_REGLAGES)) || {}); } catch (e) { /* valeurs par défaut */ }
function changerReglage(nom, valeur) {
  reglages[nom] = valeur;
  try { localStorage.setItem(CLE_REGLAGES, JSON.stringify(reglages)); } catch (e) { /* stockage bloqué */ }
}

const $ = id => document.getElementById(id);
const PLANCHE = "icones/icones.svg#";
const icone = (nom, classe = "") => `<svg class="ic ${classe}"><use href="${PLANCHE}${nom}"></use></svg>`;
// Petit identifiant unique pour chaque élément (ex. "k3f9x2a1")
const nouvelId = () => Math.random().toString(36).slice(2, 10);
// Date du jour « AAAA-MM-JJ » (heure du téléphone)
const aujourdhui = () => { const d = new Date(); return new Date(d - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10); };
// Sans accents ni majuscules : « cafe » trouve « Café »
const simplifier = t => String(t || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().trim();
// Protège le texte tapé avant de l'afficher (évite qu'un « < » casse la page)
const echapper = t => String(t ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
// N'accepte que les vrais liens web (pas de « javascript: » piégé)
const lienValide = u => /^https?:\/\/\S+$/i.test(String(u || "").trim()) ? String(u).trim() : "";
// « https://www.youtube.com/watch?v=… » -> « youtube.com »
const domaine = u => { try { return new URL(u).hostname.replace(/^www\./, ""); } catch (e) { return u; } };

const SANS_CAT = { id: "", nom: "Sans catégorie", emoji: "•" };
const categorie = id => donnees.categories.find(c => c.id === id) || SANS_CAT;

// Remet d'aplomb une astuce venue d'ailleurs (Firebase, Excel, ancienne version...)
function normaliserAstuce(id, d) {
  return {
    id,
    titre: String(d.titre || "").trim(),
    texte: String(d.texte || ""),
    cat: String(d.cat || ""),
    lien: lienValide(d.lien),
    image: lienValide(d.image),
    photo: typeof d.photo === "string" && d.photo.startsWith("data:image/") ? d.photo : "",
    favori: !!d.favori, connue: !!d.connue, vue: !!d.vue,
    cree: typeof d.cree === "number" ? d.cree : 0
  };
}
function normaliserDonnees(d) {
  const cats = Array.isArray(d.categories) ? d.categories.filter(c => c && c.id) : null;
  return {
    categories: cats ? cats.map(c => ({ id: String(c.id), nom: String(c.nom || ""), emoji: String(c.emoji || "") })) : structuredClone(CATEGORIES_DEPART),
    jour: d.jour && d.jour.date && d.jour.id ? { date: d.jour.date, id: d.jour.id } : null
  };
}

// ---------- 2. Stockage ----------
// Rangement dans Firestore :
//   users/{mon uid}/app-lpt/donnees                -> { categories, jour }
//   users/{mon uid}/app-lpt/donnees/astuces/{id}   -> une astuce par document
// Un document par astuce : une photo pèse ~100 Ko, et un document ne peut pas dépasser 1 Mo.
// Si deux appareils modifient la même astuce : la dernière modification gagne.

// Mode démo (adresse se terminant par ?demo) : données gardées dans ce navigateur
// seulement, sans connexion. Sert à tester l'appli sans compte.
const MODE_DEMO = new URLSearchParams(location.search).has("demo");

let db = null, auth = null, uid = null;
let chargeAstuces = false, chargeDonnees = false;  // les deux sont arrivés ?
const refDonnees = () => doc(db, "users", uid, NOM_APPLI, "donnees");
const refAstuce = id => doc(db, "users", uid, NOM_APPLI, "donnees", "astuces", id);

function erreurEcriture(err) {
  console.error(err);
  afficherToast("Échec de l'enregistrement", null);
}
// Ce qu'on envoie à Firestore : l'astuce sans son id (c'est déjà le nom du document)
const pourEnvoi = a => { const { id, ...reste } = a; return { ...reste, maj: serverTimestamp() }; };

// On n'attend pas la réponse du serveur : hors ligne, Firestore garde l'écriture
// en attente et l'envoie tout seul au retour du réseau.
function ecrireAstuces(liste) {
  liste.forEach(a => astuces.set(a.id, a));
  if (MODE_DEMO) { sauverDemo(); return; }
  // Par paquets de 400 (Firestore limite un paquet à 500 opérations)
  for (let i = 0; i < liste.length; i += 400) {
    const lot = writeBatch(db);
    liste.slice(i, i + 400).forEach(a => lot.set(refAstuce(a.id), pourEnvoi(a)));
    lot.commit().catch(erreurEcriture);
  }
}
const ecrireAstuce = a => ecrireAstuces([a]);

function effacerAstuces(ids) {
  ids.forEach(id => astuces.delete(id));
  if (MODE_DEMO) { sauverDemo(); return; }
  for (let i = 0; i < ids.length; i += 400) {
    const lot = writeBatch(db);
    ids.slice(i, i + 400).forEach(id => lot.delete(refAstuce(id)));
    lot.commit().catch(erreurEcriture);
  }
}

function sauverDonnees() {
  if (MODE_DEMO) { sauverDemo(); return; }
  setDoc(refDonnees(), { ...donnees, maj: serverTimestamp() }).catch(erreurEcriture);
}

function sauverDemo() {
  try { localStorage.setItem("lpt-demo", JSON.stringify({ donnees, astuces: [...astuces.values()] })); }
  catch (e) { afficherToast("Mémoire du navigateur pleine", null); }
}
function demarrerDemo() {
  try {
    const d = JSON.parse(localStorage.getItem("lpt-demo")) || {};
    donnees = normaliserDonnees(d.donnees || {});
    astuces = new Map((d.astuces || []).map(a => [a.id, normaliserAstuce(a.id, a)]));
  } catch (e) { /* on part de zéro */ }
  $("compte-info").textContent = "Mode démo : données gardées sur cet appareil seulement.";
  chargeAstuces = chargeDonnees = true;
  majSynchro();
  verifierJour();
  router();
}

function demarrerFirebase() {
  const app = initializeApp(firebaseConfig);
  auth = getAuth(app);
  // Cache local permanent : l'appli fonctionne hors ligne avec la dernière copie connue
  db = initializeFirestore(app, {
    ignoreUndefinedProperties: true,
    localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() })
  });

  onAuthStateChanged(auth, user => {
    if (!user) { montrer("connexion"); return; }
    uid = user.uid;
    $("compte-info").textContent = `Connecté : ${user.displayName || user.email}`;

    // Données générales (catégories, astuce du jour)
    onSnapshot(refDonnees(), { includeMetadataChanges: true }, snap => {
      metaDonnees = snap.metadata; majSynchro();
      if (!snap.exists()) {
        // Première utilisation : on crée les données de départ.
        // Seulement si la réponse vient vraiment du serveur, sinon on risquerait
        // d'écraser mes vraies données en étant hors ligne.
        if (snap.metadata.fromCache) return;
        donnees = normaliserDonnees({});
        sauverDonnees();
      } else {
        donnees = normaliserDonnees(snap.data());
      }
      chargeDonnees = true;
      verifierJour();
      rafraichir();
    });

    // Les astuces
    onSnapshot(collection(db, "users", uid, NOM_APPLI, "donnees", "astuces"), { includeMetadataChanges: true }, snap => {
      metaAstuces = snap.metadata; majSynchro();
      // Une copie locale vide au tout premier lancement ne veut pas dire « aucune astuce » :
      // on attend la réponse du serveur.
      if (snap.empty && snap.metadata.fromCache && !chargeAstuces) return;
      astuces = new Map(snap.docs.map(d => [d.id, normaliserAstuce(d.id, d.data())]));
      chargeAstuces = true;
      verifierJour();
      rafraichir();
    });
  });
}

async function seConnecter() {
  const fournisseur = new GoogleAuthProvider();
  $("connexion-erreur").hidden = true;
  try {
    await signInWithPopup(auth, fournisseur);
  } catch (err) {
    // Si la fenêtre de connexion est bloquée, on passe par une redirection
    if (err.code === "auth/popup-blocked" || err.code === "auth/operation-not-supported-in-this-environment") {
      signInWithRedirect(auth, fournisseur);
    } else if (err.code !== "auth/popup-closed-by-user" && err.code !== "auth/cancelled-popup-request") {
      $("connexion-erreur").textContent = "Connexion impossible : " + err.message;
      $("connexion-erreur").hidden = false;
    }
  }
}

// Indicateur de synchro (dans Réglages) : synchronisé / en attente / hors ligne
let metaDonnees = null, metaAstuces = null;
function majSynchro() {
  let classe = "ok", ic = "cloud-check", texte = "Synchronisé";
  if (MODE_DEMO) { classe = "off"; ic = "cloud-off"; texte = "Mode démo : pas de synchro"; }
  else if (!navigator.onLine) { classe = "off"; ic = "cloud-off"; texte = "Hors ligne"; }
  else if (metaDonnees?.hasPendingWrites || metaAstuces?.hasPendingWrites) { classe = "attente"; ic = "cloud-up"; texte = "En attente d'envoi"; }
  document.querySelectorAll(".synchro").forEach(el => {
    el.className = "synchro " + classe;
    el.querySelector("use").setAttribute("href", PLANCHE + ic);
    el.querySelector(".synchro-texte").textContent = texte;
  });
}
window.addEventListener("online", majSynchro);
window.addEventListener("offline", majSynchro);

// ---------- 3. Astuce du jour ----------
// Le tirage marche par « tours », comme une pioche de cartes : on tire au hasard
// parmi les astuces pas encore vues pendant ce tour. Quand toutes ont été vues,
// on remélange tout et un nouveau tour commence.
// Les astuces « Je la connais » ne sont jamais tirées.
// Le choix du jour est enregistré : le tel et le PC montrent la même astuce toute la journée.
function verifierJour() {
  if (!chargeAstuces || !chargeDonnees) return;
  const j = donnees.jour;
  if (j && j.date === aujourdhui() && astuces.has(j.id)) return; // déjà choisie aujourd'hui
  const pioche = [...astuces.values()].filter(a => !a.connue);
  if (!pioche.length) return;
  let pasVues = pioche.filter(a => !a.vue);
  if (!pasVues.length) {
    // Tour terminé : on remet tout dans la pioche
    pioche.forEach(a => { a.vue = false; });
    ecrireAstuces(pioche);
    pasVues = pioche;
  }
  const choisie = pasVues[Math.floor(Math.random() * pasVues.length)];
  choisie.vue = true;
  ecrireAstuce(choisie);
  donnees.jour = { date: aujourdhui(), id: choisie.id };
  sauverDonnees();
}
const astuceDuJour = () => {
  const j = donnees.jour;
  return j && j.date === aujourdhui() ? astuces.get(j.id) : null;
};

// ---------- 4. Navigation ----------
// Chaque écran ou fenêtre a son adresse (#reglages, #fiche=ID, #modifier=ID...) :
// ainsi le bouton « retour » du téléphone ferme la fiche ou revient en arrière,
// comme dans une vraie appli.
let ecranActuel = "";
let aNavigue = false; // a-t-on changé d'écran depuis l'ouverture ?
function aller(adresse) { aNavigue = true; location.hash = adresse; }
function retour() { if (aNavigue) history.back(); else location.replace("#"); }

function montrer(nom) {
  document.querySelectorAll(".ecran").forEach(e => { e.hidden = e.id !== "ecran-" + nom; });
  if (ecranActuel !== nom) window.scrollTo(0, 0);
  ecranActuel = nom;
}

// Les fenêtres (fiche, formulaire) s'ouvrent par-dessus l'accueil
const fenetreOuverte = () => !$("voile").hidden;
function fermerVoile() { $("voile").hidden = true; $("voile").innerHTML = ""; formEnCours = null; }

function router() {
  if (!MODE_DEMO && !uid) return; // pas encore connecté
  const h = decodeURIComponent(location.hash.slice(1));
  menuOuvert = false;
  if (h === "reglages") { fermerVoile(); afficherReglages(); montrer("reglages"); }
  else if (h === "theme") { fermerVoile(); afficherTheme(); montrer("theme"); }
  else if (h === "categories") { fermerVoile(); afficherCategories(); montrer("categories"); }
  else {
    afficherAccueil(); montrer("accueil");
    if (h.startsWith("fiche=")) {
      // (astuce introuvable, ex. juste supprimée : on ferme simplement la fenêtre)
      const a = astuces.get(h.slice(6));
      if (a) afficherFiche(a); else fermerVoile();
    } else if (h === "nouvelle" || h.startsWith("modifier=")) {
      // Le formulaire ne se redessine pas si on y est déjà (pour ne pas perdre la saisie)
      if (!formEnCours || formEnCours.adresse !== h) {
        const a = h === "nouvelle" ? null : astuces.get(h.slice(9));
        if (h !== "nouvelle" && !a) { fermerVoile(); return; }
        ouvrirFormulaire(a, h);
      }
    } else fermerVoile();
  }
  if (rechargerApres && !fenetreOuverte()) location.reload();
}
window.addEventListener("hashchange", router);

// Redessine ce qui est à l'écran quand des données arrivent (sans toucher au formulaire)
function rafraichir() {
  if (!chargeAstuces || !chargeDonnees) return;
  if (formEnCours) { afficherAccueil(); return; }
  router();
}

// ---------- 5. Affichage ----------

let filtre = "tout";      // "tout", "favoris" ou un identifiant de catégorie
let recherche = "";

// Image d'une astuce : la photo prise avec le tel en priorité, sinon l'adresse d'image
const imageDe = a => a.photo || a.image;
const htmlImage = a => imageDe(a)
  ? `<img class="image" src="${echapper(imageDe(a))}" alt="" loading="lazy" onerror="this.remove()">` : "";
const htmlLien = a => a.lien
  ? `<a class="lien" href="${echapper(a.lien)}" target="_blank" rel="noopener">${icone("external-link")}<span>${echapper(domaine(a.lien))}</span></a>` : "";
const htmlBoutons = a => `
  <div class="actions">
    <button class="btn" data-favori="${a.id}" aria-pressed="${a.favori}">${icone("star", a.favori ? "plein" : "")}Favori</button>
    <button class="btn" data-connue="${a.id}" aria-pressed="${a.connue}">${icone("check")}Je la connais</button>
  </div>`;
const meta = a => { const c = categorie(a.cat); return `${echapper(c.emoji)} ${echapper(c.nom)}`; };

function afficherAccueil() {
  // --- Carte « Astuce du jour » ---
  const a = astuceDuJour();
  if (a) {
    $("du-jour").innerHTML = `<article class="jour">
      <div class="lire" data-fiche="${a.id}">${htmlImage(a)}</div>
      <div class="corps">
        <div class="lire" data-fiche="${a.id}">
          <div class="meta">Astuce du jour · ${meta(a)}</div>
          <h2>${echapper(a.titre)}</h2>
          ${a.texte ? `<p class="texte">${echapper(a.texte)}</p>` : ""}
        </div>
        ${htmlLien(a)}
        ${htmlBoutons(a)}
      </div>
    </article>`;
  } else if (!astuces.size) {
    $("du-jour").innerHTML = `<div class="vide">${icone("bulb")}<div>Aucune astuce</div>
      <button class="btn plein" data-action="ajouter">${icone("plus")}Ajouter</button></div>`;
  } else {
    // Toutes les astuces sont marquées « Je la connais » : plus rien à tirer
    $("du-jour").innerHTML = "";
  }

  // --- Filtres : Tout, Favoris, puis une puce par catégorie ---
  // La puce choisie montre son nom ; les autres seulement leur emoji (plus compact)
  const puces = [
    { id: "tout", html: "Tout" },
    { id: "favoris", html: icone("star") + (filtre === "favoris" ? "Favoris" : ""), titre: "Favoris" },
    ...donnees.categories.map(c => ({ id: c.id, html: echapper(c.emoji) + (filtre === c.id ? " " + echapper(c.nom) : ""), titre: c.nom }))
  ];
  if (!puces.some(p => p.id === filtre)) filtre = "tout";
  $("filtres").innerHTML = puces.map(p =>
    `<button class="puce" data-filtre="${echapper(p.id)}" aria-pressed="${p.id === filtre}" ${p.titre ? `title="${echapper(p.titre)}" aria-label="${echapper(p.titre)}"` : ""}>${p.html}</button>`).join("");

  // --- Liste : les plus récentes d'abord ---
  const mots = simplifier(recherche).split(/\s+/).filter(Boolean);
  const liste = [...astuces.values()]
    .filter(x => filtre === "tout" || (filtre === "favoris" ? x.favori : x.cat === filtre))
    .filter(x => { const t = simplifier(x.titre + " " + x.texte); return mots.every(m => t.includes(m)); })
    .sort((x, y) => y.cree - x.cree);
  $("liste").innerHTML = !astuces.size ? "" : !liste.length ? `<div class="aucun">Aucun résultat</div>` :
    liste.map(x => `<button class="ligne ${x.connue ? "connue" : ""}" data-fiche="${x.id}">
      <span class="emoji">${echapper(categorie(x.cat).emoji)}</span>
      <span class="nom">${echapper(x.titre)}</span>
      ${x.favori ? icone("star", "plein fav") : ""}
      ${x.connue ? icone("check") : ""}
      ${imageDe(x) ? icone("photo") : ""}
      ${x.lien ? icone("link") : ""}
    </button>`).join("");
}

// --- Fiche d'une astuce (mini pop-up) ---
let menuOuvert = false;
function afficherFiche(a) {
  $("voile").innerHTML = `<div class="fenetre" role="dialog" aria-label="${echapper(a.titre)}">
    ${htmlImage(a)}
    <div class="corps">
      <div class="entete">
        <div class="meta">${meta(a)}</div>
        <div class="menu-wrap">
          <button class="btn-icone" data-action="menu" aria-label="Plus d'actions">${icone("dots")}</button>
          ${menuOuvert ? `<div class="menu">
            <button data-modifier="${a.id}">${icone("pencil")}Modifier</button>
            <button class="danger" data-supprimer="${a.id}">${icone("trash")}Supprimer</button>
          </div>` : ""}
        </div>
        <button class="btn-icone" data-action="fermer" aria-label="Fermer">${icone("x")}</button>
      </div>
      <h2>${echapper(a.titre)}</h2>
      ${a.texte ? `<p class="texte">${echapper(a.texte)}</p>` : ""}
      ${htmlLien(a)}
      ${htmlBoutons(a)}
    </div>
  </div>`;
  $("voile").hidden = false;
}

// --- Formulaire (ajouter / modifier) ---
// « brouillon » = l'astuce en cours de saisie ; rien n'est enregistré avant « Enregistrer ».
let formEnCours = null; // { adresse, id (null = nouvelle), brouillon, modeImage }
function ouvrirFormulaire(a, adresse, prerempli = {}) {
  const brouillon = a ? { ...a } : normaliserAstuce(null, { cat: filtreCategorie(), ...prerempli });
  formEnCours = { adresse, id: a ? a.id : null, brouillon, modeImage: !!(brouillon.image && !brouillon.photo) };
  $("voile").innerHTML = `<div class="fenetre" role="dialog" aria-label="${a ? "Modifier" : "Nouvelle astuce"}">
    <form class="form" id="form-astuce" novalidate>
      <div class="entete"><h2 style="flex:1">${a ? "Modifier" : "Nouvelle astuce"}</h2>
        <button type="button" class="btn-icone" data-action="fermer" aria-label="Fermer">${icone("x")}</button></div>
      <label class="champ"><span>Titre</span><input type="text" id="f-titre" maxlength="120" value="${echapper(brouillon.titre)}" autocomplete="off"></label>
      <p class="erreur" id="f-erreur" hidden>Il faut un titre</p>
      <div class="champ"><span>Catégorie</span><div class="cats" id="f-cats"></div></div>
      <label class="champ"><span>Texte</span><textarea id="f-texte" maxlength="3000">${echapper(brouillon.texte)}</textarea></label>
      <label class="champ"><span>Lien</span><input type="url" id="f-lien" value="${echapper(brouillon.lien)}" placeholder="https://…" autocomplete="off"></label>
      <div class="champ"><span>Image</span><div id="f-image"></div></div>
      <input type="file" id="f-fichier" accept="image/*" hidden>
      <div class="deux">
        <button type="button" class="btn" data-action="fermer">Annuler</button>
        <button type="submit" class="btn plein">Enregistrer</button>
      </div>
    </form>
  </div>`;
  $("voile").hidden = false;
  afficherCatsForm();
  afficherImageForm();
  if (!a) setTimeout(() => $("f-titre").focus(), 50);
}
// Nouvelle astuce ajoutée pendant qu'un filtre de catégorie est actif : on reprend cette catégorie
const filtreCategorie = () => donnees.categories.some(c => c.id === filtre) ? filtre : "";

function afficherCatsForm() {
  const b = formEnCours.brouillon;
  $("f-cats").innerHTML = donnees.categories.map(c =>
    `<button type="button" class="puce" data-f-cat="${echapper(c.id)}" aria-pressed="${b.cat === c.id}">${echapper(c.emoji)} ${echapper(c.nom)}</button>`).join("");
}
// Zone « Image » : aperçu (avec croix pour l'enlever), ou les deux façons d'en ajouter une
function afficherImageForm() {
  const f = formEnCours, b = f.brouillon;
  const src = b.photo || b.image;
  if (src) {
    $("f-image").innerHTML = `<div class="apercu"><img src="${echapper(src)}" alt="" onerror="this.style.opacity=.3">
      <button type="button" class="btn-icone" data-action="sans-image" aria-label="Enlever l'image">${icone("x")}</button></div>`;
  } else if (f.modeImage) {
    $("f-image").innerHTML = `<input type="url" id="f-adresse-image" placeholder="Adresse de l'image (https://…)" autocomplete="off">`;
    $("f-adresse-image").focus();
  } else {
    $("f-image").innerHTML = `<div class="deux">
      <button type="button" class="btn" data-action="photo">${icone("camera")}Photo</button>
      <button type="button" class="btn" data-action="adresse-image">${icone("link")}Adresse</button></div>`;
  }
}

// Réduit une photo à 800 px maximum et l'enregistre en JPEG (~100 Ko) :
// assez net sur un téléphone, et assez léger pour tenir dans la fiche.
async function reduirePhoto(fichier) {
  const bmp = await createImageBitmap(fichier, { imageOrientation: "from-image" });
  let taille = 800, qualite = 0.72, resultat = "";
  for (let essai = 0; essai < 6; essai++) {
    const r = Math.min(1, taille / Math.max(bmp.width, bmp.height));
    const c = document.createElement("canvas");
    c.width = Math.round(bmp.width * r); c.height = Math.round(bmp.height * r);
    c.getContext("2d").drawImage(bmp, 0, 0, c.width, c.height);
    resultat = c.toDataURL("image/jpeg", qualite);
    if (resultat.length < 180000) break;   // ~135 Ko : c'est bon
    taille *= 0.8; qualite = Math.max(0.5, qualite - 0.06);
  }
  bmp.close?.();
  return resultat;
}

// --- Réglages ---
const NOMS_THEME = { auto: "Auto", clair: "Clair", sombre: "Sombre" };
function afficherReglages() {
  $("version").textContent = VERSION;
  $("theme-actuel").textContent = NOMS_THEME[reglages.theme] || "Auto";
  document.querySelectorAll("[data-anim-choix]").forEach(b => b.setAttribute("aria-pressed", b.dataset.animChoix === reglages.animations));
  $("resume-categories").textContent = donnees.categories.map(c => c.emoji).join(" ");
  $("btn-deconnexion").hidden = MODE_DEMO;
  majSynchro();
}
function afficherTheme() {
  document.querySelectorAll("[data-theme-choix]").forEach(b => b.setAttribute("aria-pressed", b.dataset.themeChoix === reglages.theme));
}
function afficherCategories() {
  // Ne pas redessiner pendant qu'on tape dans un champ (sinon le curseur saute)
  if ($("liste-categories").contains(document.activeElement)) return;
  $("liste-categories").innerHTML = donnees.categories.map(c => `<div class="ligne-cat" data-cat="${echapper(c.id)}">
    <input class="c-emoji" value="${echapper(c.emoji)}" maxlength="8" aria-label="Emoji">
    <input class="c-nom" value="${echapper(c.nom)}" maxlength="30" aria-label="Nom">
    <button class="btn-icone" data-suppr-cat="${echapper(c.id)}" aria-label="Supprimer">${icone("trash")}</button>
  </div>`).join("");
}

// ---------- 6. Actions ----------

// Message en bas de l'écran pendant 5 secondes, avec un bouton « Annuler » optionnel
let minuteurToast = null, actionAnnuler = null;
function afficherToast(texte, annuler) {
  $("toast-texte").textContent = texte;
  $("toast-annuler").hidden = !annuler;
  actionAnnuler = annuler;
  $("toast").hidden = false;
  clearTimeout(minuteurToast);
  minuteurToast = setTimeout(() => { $("toast").hidden = true; actionAnnuler = null; }, 5000);
}
$("toast-annuler").addEventListener("click", () => {
  if (actionAnnuler) actionAnnuler();
  $("toast").hidden = true;
});

// Un seul « écouteur » pour tous les clics de la page : on regarde ce qui a été touché
document.addEventListener("click", e => {
  const el = e.target.closest("button, [data-fiche], a");
  if (!el) {
    // Toucher le fond sombre autour d'une fenêtre la ferme
    if (e.target === $("voile")) retour();
    else if (menuOuvert) { menuOuvert = false; router(); }
    return;
  }
  const d = el.dataset;

  if (d.aller) { if (d.aller === "accueil") retour(); else aller(d.aller); return; }
  if (d.fiche) { aller("fiche=" + d.fiche); return; }
  if (d.filtre) { filtre = d.filtre; afficherAccueil(); return; }

  if (d.favori || d.connue) {
    const a = astuces.get(d.favori || d.connue);
    if (!a) return;
    if (d.favori) a.favori = !a.favori; else a.connue = !a.connue;
    ecrireAstuce(a);
    rafraichir();
    return;
  }

  if (d.action === "ajouter") { aller("nouvelle"); return; }
  if (d.action === "fermer") { retour(); return; }
  if (d.action === "menu") {
    const a = astuces.get(location.hash.slice(7)); // adresse « #fiche=ID »
    if (a) { menuOuvert = !menuOuvert; afficherFiche(a); }
    return;
  }
  if (d.modifier) { location.replace("#modifier=" + d.modifier); return; }
  if (d.supprimer) {
    const a = astuces.get(d.supprimer);
    if (!a) return;
    effacerAstuces([a.id]);
    retour();   // l'accueil se redessine en revenant
    afficherToast("Astuce supprimée", () => { ecrireAstuce(a); rafraichir(); });
    return;
  }

  // Formulaire
  if (d.fCat !== undefined && formEnCours) {
    const b = formEnCours.brouillon;
    b.cat = b.cat === d.fCat ? "" : d.fCat;   // retoucher la catégorie choisie l'enlève
    afficherCatsForm();
    return;
  }
  if (d.action === "photo") { $("f-fichier").click(); return; }
  if (d.action === "adresse-image") { formEnCours.modeImage = true; afficherImageForm(); return; }
  if (d.action === "sans-image") {
    Object.assign(formEnCours.brouillon, { photo: "", image: "" });
    formEnCours.modeImage = false;
    afficherImageForm();
    return;
  }

  // Catégories (écran Réglages → Catégories)
  if (d.supprCat) {
    const avant = structuredClone(donnees.categories);
    const c = categorie(d.supprCat);
    donnees.categories = donnees.categories.filter(x => x.id !== d.supprCat);
    sauverDonnees(); afficherCategories();
    // Les astuces gardent leur catégorie : « Annuler » la fait réapparaître telle quelle
    afficherToast(`« ${c.nom} » supprimée`, () => { donnees.categories = avant; sauverDonnees(); router(); });
    return;
  }
});

$("btn-ajouter").addEventListener("click", () => aller("nouvelle"));
$("btn-connexion").addEventListener("click", seConnecter);
$("btn-deconnexion").addEventListener("click", () => {
  if (confirm("Se déconnecter ? Tes astuces restent en sécurité dans ton compte.")) signOut(auth).then(() => location.reload());
});
$("recherche").addEventListener("input", e => { recherche = e.target.value; afficherAccueil(); });

// Photo choisie (ou prise) : on la réduit avant de la garder
document.addEventListener("change", async e => {
  if (e.target.id !== "f-fichier" || !formEnCours) return;
  const fichier = e.target.files[0];
  e.target.value = "";
  if (!fichier) return;
  try {
    formEnCours.brouillon.photo = await reduirePhoto(fichier);
    formEnCours.brouillon.image = "";
    afficherImageForm();
  } catch (err) {
    afficherToast("Photo illisible", null);
  }
});
// Adresse d'image tapée : on l'affiche dès qu'on quitte le champ
document.addEventListener("focusout", e => {
  if (e.target.id === "f-adresse-image" && formEnCours) {
    const u = lienValide(e.target.value);
    if (u) { formEnCours.brouillon.image = u; formEnCours.modeImage = false; afficherImageForm(); }
  }
  // Catégories : enregistrées dès qu'on quitte le champ
  const ligne = e.target.closest?.(".ligne-cat");
  if (ligne) {
    const c = donnees.categories.find(x => x.id === ligne.dataset.cat);
    if (!c) return;
    const emoji = ligne.querySelector(".c-emoji").value.trim(), nom = ligne.querySelector(".c-nom").value.trim();
    if (nom && (nom !== c.nom || emoji !== c.emoji)) { c.nom = nom; c.emoji = emoji; sauverDonnees(); }
  }
});
document.addEventListener("input", e => { if (e.target.id === "f-titre") $("f-erreur").hidden = true; });

// Enregistrer le formulaire
document.addEventListener("submit", e => {
  if (e.target.id !== "form-astuce") return;
  e.preventDefault();
  const f = formEnCours, b = f.brouillon;
  b.titre = $("f-titre").value.trim();
  if (!b.titre) { $("f-erreur").hidden = false; $("f-titre").focus(); return; }
  b.texte = $("f-texte").value.trim();
  b.lien = lienValide($("f-lien").value);
  if ($("f-adresse-image")) b.image = lienValide($("f-adresse-image").value) || b.image;
  if (!f.id) { b.id = nouvelId(); b.cree = Date.now(); }
  ecrireAstuce(normaliserAstuce(b.id, b));
  formEnCours = null;
  // Après un ajout, on revient à l'accueil ; après une modif, à la fiche
  if (f.id) location.replace("#fiche=" + f.id); else retour();
  verifierJour(); // la toute première astuce devient l'astuce du jour
  rafraichir();
});

// Thème et animations (réglages de cet appareil)
$("choix-theme").addEventListener("click", e => {
  const b = e.target.closest("[data-theme-choix]");
  if (!b) return;
  changerReglage("theme", b.dataset.themeChoix);
  window.appliquerTheme(b.dataset.themeChoix);
  afficherTheme();
});
function appliquerAnimations() {
  document.documentElement.dataset.anim = reglages.animations;
}
appliquerAnimations();
$("choix-animations").addEventListener("click", e => {
  const b = e.target.closest("[data-anim-choix]");
  if (!b) return;
  changerReglage("animations", b.dataset.animChoix);
  appliquerAnimations();
  afficherReglages();
});

$("btn-ajout-categorie").addEventListener("click", () => {
  const c = { id: nouvelId(), nom: "Nouvelle", emoji: "📌" };
  donnees.categories.push(c);
  sauverDonnees();
  afficherCategories();
  const champ = document.querySelector(`[data-cat="${c.id}"] .c-nom`);
  champ.focus(); champ.select();
});

// Lien reçu via le menu « Partager » d'Android (…/?title=…&text=…&url=…) :
// on ouvre une nouvelle fiche pré-remplie.
function lirePartage() {
  const p = new URLSearchParams(location.search);
  if (!p.has("url") && !p.has("text") && !p.has("title")) return;
  const texte = p.get("text") || "";
  // Selon l'appli d'origine, le lien arrive dans « url » ou au milieu du texte
  const lien = lienValide(p.get("url")) || (texte.match(/https?:\/\/\S+/) || [""])[0];
  const titre = (p.get("title") || texte.replace(lien, "")).trim().slice(0, 120);
  history.replaceState(null, "", location.pathname + (MODE_DEMO ? "?demo" : ""));
  const attendre = setInterval(() => {
    if (!chargeAstuces || !chargeDonnees) return;
    clearInterval(attendre);
    aNavigue = true;
    history.pushState(null, "", "#nouvelle");
    ouvrirFormulaire(null, "nouvelle", { titre, lien });
  }, 200);
}

// ---------- 7. Export / import Excel ----------
// Fichier : feuille « Astuces » (une ligne par astuce) et feuille « Catégories ».
// Les photos prises avec le tel ne vont pas dans Excel (trop lourdes pour une case) :
// à l'import, une astuce qui porte le même titre garde sa photo.
const COLONNES = [
  { header: "Catégorie", key: "cat", width: 16 },
  { header: "Titre", key: "titre", width: 40 },
  { header: "Texte", key: "texte", width: 70 },
  { header: "Lien", key: "lien", width: 30 },
  { header: "Image", key: "image", width: 30 },
  { header: "Favori", key: "favori", width: 9 },
  { header: "Connue", key: "connue", width: 9 },
  { header: "Ajoutée le", key: "ajoutee", width: 12 }
];
const ouiNon = v => v ? "Oui" : "";
const dateTexte = ms => ms ? new Date(ms - new Date(ms).getTimezoneOffset() * 60000).toISOString().slice(0, 10) : "";

$("btn-export").addEventListener("click", async () => {
  try {
    const ExcelJS = await chargerExcelJS();
    const classeur = new ExcelJS.Workbook();
    classeur.creator = "LPT";
    const ordreCat = id => { const i = donnees.categories.findIndex(c => c.id === id); return i < 0 ? 999 : i; };
    const lignes = [...astuces.values()]
      .sort((a, b) => ordreCat(a.cat) - ordreCat(b.cat) || a.cree - b.cree)
      .map(a => ({ cat: categorie(a.cat).id ? categorie(a.cat).nom : "", titre: a.titre, texte: a.texte, lien: a.lien,
        image: a.image, favori: ouiNon(a.favori), connue: ouiNon(a.connue), ajoutee: dateTexte(a.cree) }));
    const ws = ajouterFeuille(classeur, "Astuces", COLONNES, lignes);
    ws.getColumn("texte").alignment = { wrapText: true, vertical: "top" };
    ajouterFeuille(classeur, "Catégories", [
      { header: "Emoji", key: "emoji", width: 8 },
      { header: "Catégorie", key: "nom", width: 20 }
    ], donnees.categories.map(c => ({ emoji: c.emoji, nom: c.nom })));
    await telechargerClasseur(classeur, NOM_APPLI);
  } catch (err) {
    afficherToast("Export impossible : " + err.message, null);
  }
});

$("btn-import").addEventListener("click", () => $("fichier-import").click());
$("fichier-import").addEventListener("change", async e => {
  const fichier = e.target.files[0];
  e.target.value = ""; // pour pouvoir réimporter le même fichier plus tard
  if (!fichier) return;
  if (!confirm("Remplacer toutes tes astuces par le contenu de ce fichier ?")) return;
  try {
    const classeur = await lireClasseur(fichier);
    const lignes = lireFeuille(classeur.getWorksheet("Astuces")).filter(l => l["Titre"]);
    const lignesCats = lireFeuille(classeur.getWorksheet("Catégories")).filter(l => l["Catégorie"]);
    if (!lignes.length) throw new Error("feuille « Astuces » introuvable ou vide");

    // Catégories : on garde les identifiants existants quand le nom correspond
    const categories = [];
    const trouverCat = (nom, emoji) => {
      if (!nom) return "";
      let c = categories.find(x => simplifier(x.nom) === simplifier(nom));
      if (!c) {
        const ancienne = donnees.categories.find(x => simplifier(x.nom) === simplifier(nom));
        c = { id: ancienne ? ancienne.id : nouvelId(), nom, emoji: emoji || ancienne?.emoji || "📌" };
        categories.push(c);
      } else if (emoji) c.emoji = emoji;
      return c.id;
    };
    lignesCats.forEach(l => trouverCat(l["Catégorie"], l["Emoji"]));

    // Astuces : celles qui portent le même titre gardent leur identifiant, leur photo et leur tirage
    const parTitre = new Map([...astuces.values()].map(a => [simplifier(a.titre), a]));
    const oui = v => simplifier(v) === "oui";
    const maintenant = Date.now();
    const nouvelles = lignes.map((l, i) => {
      const ancienne = parTitre.get(simplifier(l["Titre"]));
      const d = l["Ajoutée le"] ? Date.parse(l["Ajoutée le"]) : NaN;
      return normaliserAstuce(ancienne ? ancienne.id : nouvelId(), {
        titre: l["Titre"], texte: l["Texte"], cat: trouverCat(l["Catégorie"]),
        lien: l["Lien"], image: l["Image"], photo: ancienne?.photo, vue: ancienne?.vue,
        favori: oui(l["Favori"]), connue: oui(l["Connue"]),
        cree: ancienne?.cree || (isNaN(d) ? maintenant - i : d)   // garde l'ordre du fichier
      });
    });

    // Photo de l'état d'avant, pour « Annuler »
    const avantAstuces = [...astuces.values()], avantDonnees = structuredClone(donnees);
    const garder = new Set(nouvelles.map(a => a.id));
    effacerAstuces(avantAstuces.filter(a => !garder.has(a.id)).map(a => a.id));
    ecrireAstuces(nouvelles);
    donnees.categories = categories.length ? categories : donnees.categories;
    sauverDonnees();
    verifierJour();
    router();
    afficherToast(`Import terminé : ${nouvelles.length} astuces`, () => {
      const remises = new Set(avantAstuces.map(a => a.id));
      effacerAstuces([...astuces.keys()].filter(id => !remises.has(id)));
      ecrireAstuces(avantAstuces);
      donnees = avantDonnees;
      sauverDonnees();
      router();
    });
  } catch (err) {
    afficherToast("Import impossible : " + err.message, null);
  }
});

// ---------- 8. Démarrage ----------

// Service worker : permet d'ouvrir l'appli sans réseau, de l'installer, et de la mettre à jour
let rechargerApres = false;
if ("serviceWorker" in navigator && (location.protocol === "https:" || location.hostname === "localhost")) {
  // updateViaCache "none" : pour savoir s'il existe une nouvelle version, toujours
  // demander sw.js au serveur, jamais à la mémoire du téléphone
  navigator.serviceWorker.register("sw.js", { updateViaCache: "none" }).then(reg => {
    // Sur Android, l'appli installée reste souvent ouverte en arrière-plan :
    // à chaque retour dessus, on vérifie s'il y a une mise à jour
    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "visible") reg.update().catch(() => {});
    });
  }).catch(() => {});
  // Quand la nouvelle version prend le relais, on recharge la page une fois pour l'afficher
  // (pas à la toute première installation ; et on attend qu'aucune fiche ne soit ouverte)
  const avaitUneVersion = !!navigator.serviceWorker.controller;
  navigator.serviceWorker.addEventListener("controllerchange", () => {
    if (!avaitUneVersion || rechargerApres) return;
    rechargerApres = true;
    if (!fenetreOuverte()) location.reload();
  });
}

// En revenant sur l'appli le lendemain (elle reste souvent ouverte en arrière-plan) :
// nouvelle astuce du jour
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible") { verifierJour(); rafraichir(); }
});

$("version").textContent = VERSION;
lirePartage();
if (MODE_DEMO) demarrerDemo(); else demarrerFirebase();
