/* Générateur de fiches navette NOF 27, cœur commun au navigateur et aux tests.
   Aucun renseignement sur la formation ne figure ici : tout vient du classeur. */
(function (global) {
  'use strict';

  var W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  var W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

  var FEUILLES_REQUISES = ['UE', 'EC', 'Competences_UE', 'MCCC', 'Seances',
    'Annexe_organisation_anterieure', 'Ref_Textes', 'Ref_Niveaux',
    'Ref_Blocs_maquette', 'Ref_Gabarit_fiche'];

  /* ------------------------------------------------------------ utilitaires */
  function val(x) { return (x === null || x === undefined) ? '' : String(x).trim(); }

  function parser() {
    return new global.DOMParser();
  }
  function serialiser(doc) {
    return new global.XMLSerializer().serializeToString(doc);
  }
  function els(noeud, ns, nom) {
    return Array.prototype.slice.call(noeud.getElementsByTagNameNS(ns, nom));
  }

  /* ---------------------------------------------------------------- la base */
  function lireBase(classeur) {
    var manquantes = FEUILLES_REQUISES.filter(function (n) {
      return classeur.SheetNames.indexOf(n) < 0;
    });
    if (manquantes.length) {
      throw new Error("Ce classeur n'est pas une base de fiches navette. "
        + 'Feuilles absentes : ' + manquantes.join(', ') + '.');
    }
    var b = { feuilles: {} };
    classeur.SheetNames.forEach(function (n) {
      b.feuilles[n] = global.XLSX.utils.sheet_to_json(classeur.Sheets[n], { defval: '' });
    });
    b.textes = {};
    b.feuilles.Ref_Textes.forEach(function (l) {
      if (val(l['Clé'])) b.textes[val(l['Clé'])] = String(l.Texte);
    });
    b.niveaux = {};
    b.feuilles.Ref_Niveaux.forEach(function (l) { b.niveaux[val(l.Code)] = l; });
    b.blocs = {};
    b.feuilles.Ref_Blocs_maquette.forEach(function (l) { b.blocs[val(l.Bloc)] = l; });
    b.regimes = {};
    b.feuilles.Ref_Gabarit_fiche.forEach(function (l) { b.regimes[val(l['Régime'])] = l; });
    b.texte = function (cle, valeurs) {
      if (!(cle in b.textes)) {
        throw new Error('La clé de texte « ' + cle + ' » est absente de la feuille Ref_Textes.');
      }
      return b.textes[cle].replace(/\{(\w+)\}/g, function (tout, nom) {
        if (!valeurs || !(nom in valeurs)) {
          throw new Error('Le texte « ' + cle + ' » emploie la variable ' + nom
            + ', qui n\'est pas fournie.');
        }
        return valeurs[nom];
      });
    };
    b.grouper = function (feuille, cle) {
      var m = {};
      b.feuilles[feuille].forEach(function (l) {
        var k = val(l[cle]);
        (m[k] = m[k] || []).push(l);
      });
      return m;
    };
    return b;
  }

  /* ------------------------------------------------------- jetons du modèle */
  function noeudsTexte(racine) { return els(racine, W, 't'); }

  function trouverJeton(racine, jeton) {
    var n = noeudsTexte(racine);
    for (var i = 0; i < n.length; i++) {
      if ((n[i].textContent || '').indexOf(jeton) >= 0) return n[i];
    }
    return null;
  }

  function paragrapheDe(noeud) {
    var p = noeud;
    while (p && !(p.namespaceURI === W && p.localName === 'p')) p = p.parentNode;
    return p;
  }
  function ligneDe(noeud) {
    var p = noeud;
    while (p && !(p.namespaceURI === W && p.localName === 'tr')) p = p.parentNode;
    return p;
  }

  /* Remplace un jeton par une valeur. Un tableau produit autant de paragraphes. */
  function remplir(racine, jeton, valeur) {
    var cible = trouverJeton(racine, jeton);
    if (!cible) return false;
    var lignes = Array.isArray(valeur) ? valeur.slice() : [String(valeur)];
    if (!lignes.length) lignes = [''];
    var para = paragrapheDe(cible);
    cible.textContent = (cible.textContent || '').replace(jeton, lignes[0]);
    if (para && lignes.length > 1) {
      var apres = para;
      for (var i = 1; i < lignes.length; i++) {
        var clone = para.cloneNode(true);
        var t = noeudsTexte(clone);
        if (t.length) {
          t[0].textContent = lignes[i];
          for (var k = 1; k < t.length; k++) t[k].textContent = '';
        }
        if (apres.nextSibling) apres.parentNode.insertBefore(clone, apres.nextSibling);
        else apres.parentNode.appendChild(clone);
        apres = clone;
      }
    }
    return true;
  }

  /* Duplique la ligne de tableau qui porte un jeton, une fois par donnée. */
  function remplirLignes(racine, jetonAncre, donnees, remplirUne) {
    var cible = trouverJeton(racine, jetonAncre);
    if (!cible) return;
    var modele = ligneDe(cible);
    if (!modele) return;
    var parent = modele.parentNode;
    donnees.forEach(function (d) {
      var clone = modele.cloneNode(true);
      remplirUne(clone, d);
      noeudsTexte(clone).forEach(function (t) {
        t.textContent = (t.textContent || '').replace(/§[A-Z0-9]+§/g, '');
      });
      parent.insertBefore(clone, modele);
    });
    parent.removeChild(modele);
  }

  function cocher(racine, coche) {
    els(racine, W14, 'checkbox').forEach(function (cb) {
      els(cb, W14, 'checked').forEach(function (e) {
        e.setAttributeNS(W14, 'w14:val', coche ? '1' : '0');
      });
      var sdt = cb.parentNode && cb.parentNode.parentNode;
      if (sdt) {
        noeudsTexte(sdt).forEach(function (t) { t.textContent = coche ? '☒' : '☐'; });
      }
    });
  }

  function idsUniques(doc) {
    var n = 900000;
    els(doc, W, 'sdtPr').forEach(function (pr) {
      els(pr, W, 'id').forEach(function (e) { n += 1; e.setAttributeNS(W, 'w:val', String(n)); });
    });
  }

  function nettoyerJetonsRestants(doc) {
    noeudsTexte(doc).forEach(function (t) {
      t.textContent = (t.textContent || '').replace(/§[A-Z0-9]+§/g, '');
    });
  }

  /* ------------------------------------------------------------- l'annexe */
  function ajouterAnnexe(doc, b, annexe) {
    var corps = els(doc, W, 'body')[0];
    var sectPr = els(corps, W, 'sectPr')[0];
    function p(texte, gras, italique, saut) {
      var xml = '<w:p xmlns:w="' + W + '"><w:r>'
        + (saut ? '<w:br w:type="page"/>' : '')
        + '<w:rPr>' + (gras ? '<w:b/>' : '') + (italique ? '<w:i/>' : '') + '</w:rPr>'
        + '<w:t xml:space="preserve"></w:t></w:r></w:p>';
      var noeud = parser().parseFromString(xml, 'application/xml').documentElement;
      noeud = doc.importNode ? doc.importNode(noeud, true) : noeud;
      els(noeud, W, 't')[0].textContent = texte;
      if (sectPr) corps.insertBefore(noeud, sectPr); else corps.appendChild(noeud);
    }
    p('', false, false, true);
    p(b.texte('annexe_titre', {}), true, false, false);
    p(b.texte('annexe_avertissement', {}), false, true, false);
    p(b.texte('annexe_ecart', { ecart: val(annexe['Écart de volumes']) }), true, false, false);
    p(val(annexe["Organisation antérieure, texte de l annexe"]), false, false, false);
  }

  /* --------------------------------------------------------- une fiche UE */
  function genererFiche(modeleBuf, b, ue, ecs, comps, mccc, seances, annexe) {
    var zipModele;
    return global.JSZip.loadAsync(modeleBuf).then(function (z) {
      zipModele = z;
      return zipModele.file('word/document.xml').async('string');
    }).then(function (xml) {
      var doc = parser().parseFromString(xml, 'application/xml');
      var ident = val(ue.ID_UE), libelle = val(ue['Libellé NOF 27']);
      var niv = val(ue['Niveau visé']), sem = val(ue.Semestre), bloc = val(ue.Bloc);
      if (!b.niveaux[niv]) throw new Error('Niveau « ' + niv + ' » absent de Ref_Niveaux (UE ' + ident + ').');
      if (!b.blocs[bloc]) throw new Error('Bloc « ' + bloc + ' » absent de Ref_Blocs_maquette (UE ' + ident + ').');
      var n = b.niveaux[niv];
      var entete = b.texte('entete_fiche', { id: ident, libelle: libelle, niveau: niv, semestre: sem });

      remplir(doc, '§ENTETE§', entete);
      remplir(doc, '§LIBELLE§', b.texte('libelle_ue', { libelle: libelle }));
      var resp = val(ue['Responsable proposé']);
      remplir(doc, '§RESPONSABLE§', resp ? b.texte('responsable_propose', { responsable: resp })
        : b.texte('responsable_invite', {}));
      remplir(doc, '§SEMESTRE§', sem);
      remplir(doc, '§NIVEAU§', niv + ' ' + val(n.Qualificatif));
      remplir(doc, '§BLOC§', val(b.blocs[bloc]['Libellé pour la fiche navette']));
      remplir(doc, '§ECTS§', val(ue.ECTS));
      remplir(doc, '§VOLETU§', b.texte('volume_etudiant', { volume: val(ue['Vol. étudiant (h)']) }));
      remplir(doc, '§VOLUMES§', b.texte('volumes_cm_td_tp',
        { cm: val(ue['CM (h)']), td: val(ue['TD (h)']), tp: val(ue['TP (h)']) }));
      var anc = val(ue['Ancienne UE']);
      remplir(doc, '§ANCIENNE§', anc ? anc : b.texte('ancienne_ue_absente', {}));
      var lignesEc = ecs.map(function (e) {
        return b.texte('ec_ligne', { libelle: val(e['Libellé de l EC']), heures: val(e['Total (h)']) });
      });
      remplir(doc, '§EC§', lignesEc.length ? lignesEc : [b.texte('ec_absent', {})]);
      remplir(doc, '§MAQUETTE§', val(ue['Votre UE par rapport à la maquette actuelle']));
      remplir(doc, '§AMORCE§', b.texte('amorce_competences', {
        niveau_libelle: val(n['Libellé du niveau']), qualificatif: val(n.Qualificatif), semestre: sem
      }));
      remplir(doc, '§ENTETEIND§', b.texte('entete_indicateur', { niveau: niv }));

      remplirLignes(doc, '§CCODE§', comps.length ? comps : [null], function (tr, c) {
        if (!c) return;
        remplir(tr, '§CCODE§', val(c.Code_competence));
        remplir(tr, '§CLIB§', val(c['Libellé de la composante (RNCP 38185)']));
        var ind = val(c['Indicateurs du comité au niveau visé']);
        remplir(tr, '§CIND§', ind ? ind.split('\n').filter(function (x) { return x.trim(); })
          : b.texte('indicateur_absent', {}));
      });

      remplirLignes(doc, '§M1§', mccc.length ? [mccc] : [null, null, null], function (tr, lot) {
          if (!lot) return;
          remplir(tr, '§M1§', val(lot[0]['Type de contrôle']));
          var pct = lot[0]['Pourcentage EvC'];
          remplir(tr, '§M2§', (pct === '' || pct === 0 || pct === null) ? '' : pct + ' %');
          var parCol = {};
          lot.forEach(function (e) {
            var g = b.regimes[val(e['Régime'])];
            if (!g) return;
            var nb = e['Nombre d épreuves'];
            parCol[g["Colonne Nb d'épreuves"]] = (nb === '' || nb === null) ? '' : String(parseInt(nb, 10));
            parCol[g['Colonne Nature']] = val(e['Nature de l épreuve']);
            parCol[g['Colonne Durée']] = val(e['Durée de l épreuve']);
          });
          for (var j = 3; j <= 11; j++) remplir(tr, '§M' + j + '§', parCol[j] || '');
      });

      var rappels = mccc.map(function (e) {
        var m = val(e['Modalité décrite dans la fiche 2026-27']);
        return m ? b.texte('mccc_rappel_ligne', { regime: val(e['Régime']), modalite: m }) : '';
      }).filter(function (x) { return x; });
      remplir(doc, '§MRAPPEL§', rappels.length
        ? [b.texte('mccc_rappel_titre', {})].concat(rappels) : '');

      var lignesSeances = seances.length
        ? seances.map(function (s, i) { return { s: s, ordre: i + 1 }; })
          .concat([1, 2, 3].map(function (k) { return { s: null, ordre: seances.length + k }; }))
        : Array.apply(null, { length: 11 }).map(function (x, i) { return { s: null, ordre: i + 1 }; });
      remplirLignes(doc, '§SORDRE§', lignesSeances, function (tr, d) {
        remplir(tr, '§SORDRE§', String(d.ordre));
        var s = d.s;
        remplir(tr, '§STYPE§', s ? val(s.Type) : '');
        remplir(tr, '§SINTERV§', s ? val(s.Intervenant) : '');
        remplir(tr, '§SINST§', s ? val(s.Institution) : '');
        remplir(tr, '§SMAIL§', s ? val(s.Courriel) : '');
        remplir(tr, '§SSALLE§', s ? val(s['Salle souhaitée ou lieu']) : '');
        remplir(tr, '§SMAT§', s ? val(s['Matériel particulier']) : '');
        remplir(tr, '§SINTIT§', s ? val(s['Intitulé de la séance']) : '');
        cocher(tr, !!(s && val(s['Extérieur'])));
      });

      var indispo = val(ue['Indisponibilités des intervenants']);
      remplir(doc, '§INDISPO§', indispo ? b.texte('indispo_prefixe', { indisponibilites: indispo }) : '');

      if (annexe) ajouterAnnexe(doc, b, annexe);
      nettoyerJetonsRestants(doc);
      idsUniques(doc);

      var sortie = zipModele;
      sortie.file('word/document.xml', serialiser(doc));

      var entetes = Object.keys(zipModele.files).filter(function (f) {
        return /^word\/header\d*\.xml$/.test(f);
      });
      return Promise.all(entetes.map(function (f) {
        return zipModele.file(f).async('string').then(function (h) {
          sortie.file(f, h.replace(/NIVEAU\s*[123]/g, 'NIVEAU ' + niv.charAt(1))
            .replace(/\bN[123]\b/g, niv));
        });
      })).then(function () {
        return zipModele.file('docProps/core.xml').async('string');
      }).then(function (c) {
        c = c.replace(/<dc:title>[\s\S]*?<\/dc:title>/, '<dc:title>' + echapper(entete) + '</dc:title>')
          .replace(/<cp:category>[\s\S]*?<\/cp:category>/, '<cp:category>' + niv + '</cp:category>');
        if (c.indexOf('<cp:category>') < 0) {
          c = c.replace('</cp:coreProperties>', '<cp:category>' + niv + '</cp:category></cp:coreProperties>');
        }
        sortie.file('docProps/core.xml', c);
        return sortie.generateAsync({ type: global.NOF27_TYPE_SORTIE || 'uint8array' })
          .then(function (data) {
            return { nom: nomFichier(b.texte('nom_fichier',
              { id: ident, libelle: libelle, niveau: niv })), data: data };
          });
      });
    });
  }

  function echapper(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  function nomFichier(t) {
    return String(t).replace(/[^\p{L}\p{N}\s\-',.]/gu, '').replace(/\s+/g, ' ').trim();
  }

  /* ------------------------------------------------------- toutes les fiches */
  function genererTout(classeur, modeleBuf, progression) {
    var b = lireBase(classeur);
    return global.JSZip.loadAsync(modeleBuf).then(function (zipControle) {
      if (!zipControle.file('word/document.xml')) {
        throw new Error("Ce fichier n'est pas un document Word valide.");
      }
      var ecs = b.grouper('EC', 'ID_UE'), comps = b.grouper('Competences_UE', 'ID_UE');
      var mccc = b.grouper('MCCC', 'ID_UE'), seances = b.grouper('Seances', 'ID_UE');
      var annexes = {};
      b.feuilles.Annexe_organisation_anterieure.forEach(function (l) {
        annexes[val(l.ID_UE)] = l;
      });
      var ues = b.feuilles.UE.filter(function (u) { return val(u.ID_UE); });
      var resultats = [];
      return ues.reduce(function (chaine, ue, i) {
        return chaine.then(function () {
          var id = val(ue.ID_UE);
          var ss = (seances[id] || []).slice().sort(function (a, c) {
            return (parseInt(a.Ordre, 10) || 0) - (parseInt(c.Ordre, 10) || 0);
          });
          return genererFiche(modeleBuf, b, ue, ecs[id] || [], comps[id] || [],
            mccc[id] || [], ss, annexes[id]).then(function (f) {
            resultats.push({ id: id, nom: f.nom, data: f.data,
              regime: val(ue['Régime de préremplissage']),
              competences: (comps[id] || []).length, seances: ss.length,
              epreuves: (mccc[id] || []).length });
            if (progression) progression(i + 1, ues.length, id);
          });
        });
      }, Promise.resolve()).then(function () { return resultats; });
    });
  }

  global.NOF27 = { lireBase: lireBase, genererTout: genererTout,
                   FEUILLES_REQUISES: FEUILLES_REQUISES };
})(typeof window !== 'undefined' ? window : globalThis);
