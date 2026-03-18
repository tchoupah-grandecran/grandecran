// V1602 - AVATAR GALLERY IMPLEMENTATION
// THIS CODE HAS BEEN WRITTEN BY KEVIN LEMESLE, WITH THE HELP OF CLAUDE.AI

// ==========================================
// CONFIGURATION
// ==========================================
const SPREADSHEET_ID = '16qraS8SqOar-t97lBKvC3A9enlz_DtRgZ6SXOmGzXBw';
const USER_TRACKING_SHEET_ID = '1YhZee0e7T_5f8M_F3GuUu81i9tIp77gWLwPe-0bMqbQ';
const TMDB_API_KEY = PropertiesService.getScriptProperties().getProperty('TMDB_API_KEY');

// ==========================================
// PRESET AVATAR GALLERY
// ==========================================
const AVATAR_PRESETS = [
  'https://i.imgur.com/54i18a4.png',  // Kevin Lemesle
  'https://i.imgur.com/wh92836.png',  // Olivia Colman
  'https://i.imgur.com/0OmLvJA.png',  // Kit Connor
  'https://i.imgur.com/6GdXcue.png',  // Paul Mescal
  'https://i.imgur.com/gtbDH4p.png',  // Pedro Pascal
  'https://i.imgur.com/0m6rNRf.png',  // Meryl Streep
  'https://i.imgur.com/NWaeMDI.png',  // Camille Cottin
  'https://i.imgur.com/PYzEx97.png',  // Bastien Bouillon
  'https://i.imgur.com/V5RJuqj.png',  // Laure Calamy
  'https://i.imgur.com/884g6CY.png',  // Virginie Efira
  'https://i.imgur.com/cNIOilm.png',  // Maggie Smith
  'https://i.imgur.com/xEJPFzP.png',  // Audrey Fleurot
  'https://i.imgur.com/SRm2Lvv.png'   // Pio Marmaï
];

function getAvatarPresets() {
  return AVATAR_PRESETS;
}

// ==========================================
// WEB APP — POINT D'ENTRÉE UNIQUE
// ==========================================

function getUserEmail() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (e) {
    Logger.log('Error getting user email: ' + e);
    return null;
  }
}

function doGet(e) {
  try {
    // ── Mode ping léger : vérifie si une session est active ──────────────
    // Appelé par le wrapper PWA via ?ping=1 pour tester l'auth sans charger toute l'app
    if (e && e.parameter && e.parameter.ping === '1') {
      try {
        const email = Session.getActiveUser().getEmail();
        if (email) {
          return ContentService
            .createTextOutput(JSON.stringify({ ok: true, email: email }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      } catch (err) {}
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── Tentative de récupération de l'email utilisateur ─────────────────
    let userEmail = null;
    try {
      userEmail = Session.getActiveUser().getEmail();
    } catch (err) {
      Logger.log('Impossible de récupérer l\'email: ' + err);
    }

    // ── Pas de session : retourner une page qui signale GAS_NEED_AUTH ─────
    // Ne jamais retourner un message d'erreur texte — toujours du HTML
    // pour que le wrapper PWA puisse recevoir le postMessage
    if (!userEmail) {
      return HtmlService.createHtmlOutput(`
        <!DOCTYPE html><html><head>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        </head><body style="margin:0;background:#000;">
        <script>
          try {
            window.parent.postMessage({ type: 'GAS_NEED_AUTH' }, '*');
          } catch(e) {}
        <\/script>
        </body></html>
      `)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // ── Profil utilisateur ────────────────────────────────────────────────
    const profile = getUserProfile();
    const isProfileValid = profile &&
                           typeof profile === 'object' &&
                           profile.firstName &&
                           profile.spreadsheetId;

    // ── Profil invalide ou inexistant : page d'inscription ────────────────
    if (!isProfileValid) {
      Logger.log('doGet: Profil invalide ou inexistant, redirection inscription');
      PropertiesService.getUserProperties().deleteAllProperties();

      return HtmlService.createTemplateFromFile('registration')
        .evaluate()
        .setTitle('Inscription - Grand Ecran')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover')
        .addMetaTag('apple-mobile-web-app-capable', 'yes')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // ── Profil valide : charger l'application ─────────────────────────────
    const t = HtmlService.createTemplateFromFile('index');
    t.userFirstName  = profile.firstName;
    t.userEmail      = profile.email;
    t.userPreferences = JSON.stringify(profile.preferences);
    t.userAvatar     = profile.avatar || AVATAR_PRESETS[0];

    return t.evaluate()
      .setTitle('Grand Ecran')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover')
      .addMetaTag('apple-mobile-web-app-capable', 'yes')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (e) {
    Logger.log('doGet error: ' + e.toString());
    // Même en cas d'erreur critique : retourner du HTML avec GAS_NEED_AUTH
    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html><html><body style="margin:0;background:#000;">
      <script>
        try {
          window.parent.postMessage({ type: 'GAS_NEED_AUTH' }, '*');
        } catch(e) {}
      <\/script>
      </body></html>
    `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// FONCTION PRINCIPALE — Films + Genres
// ==========================================
function getFilmsEtGenres() {
  try {
    const spreadsheetId = getUserSpreadsheetId();
    const films  = getFilmsANoter(spreadsheetId);
    const genres = getGenres(spreadsheetId);
    return { films: films, genres: genres };
  } catch (e) {
    Logger.log('Erreur getFilmsEtGenres : ' + e);
    return { films: [], genres: getDefaultGenres() };
  }
}

// ==========================================
// PARSING EMAIL PATHÉ
// ==========================================
function parsePatheEmail(message) {
  if (!message) {
    Logger.log('Message undefined');
    return null;
  }

  const subject = message.getSubject();
  if (!subject.includes('Confirmation de commande les cinémas Pathé')) {
    Logger.log('Pas un email de confirmation Pathé');
    return null;
  }

  const bodyData = parseEmailBody(message);
  if (bodyData && bodyData.titre) {
    bodyData.titre = decodeHtmlEntities(bodyData.titre);
    return bodyData;
  }

  return null;
}

function parseEmailBody(message) {
  try {
    const htmlBody  = message.getBody();
    const plainBody = message.getPlainBody();
    const data = {};

    // ── Titre ──────────────────────────────────────────────────────────────
    let m = htmlBody.match(/<h1[^>]*>\s*<span[^>]*><\/span>\s*([^<]+)\s*<\/h1>/i);
    if (!m) m = htmlBody.match(/<h1[^>]*>\s*([^<]+)\s*<\/h1>/i);
    if (!m) m = plainBody.match(/Film\s*:?\s*([^\n\r]+)/i);
    if (!m) return null;

    let titre = m[1].replace(/\s+/g, ' ').trim();
    titre = titre.replace(/^La Soirée des Passionnés\s*:\s*/i, '').trim();
    data.titre = titre;

    // ── Date & heure ────────────────────────────────────────────────────────
    m = htmlBody.match(/(Lundi|Mardi|Mercredi|Jeudi|Vendredi|Samedi|Dimanche)?\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})[,\s]*(\d{1,2}:\d{2})/i);
    if (m) {
      data.date  = formatDate(m[2]);
      data.heure = m[3];
    } else {
      m = htmlBody.match(/(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/);
      if (m) data.date = formatDate(m[1]);
      m = htmlBody.match(/(\d{1,2}:\d{2})/);
      if (m) data.heure = m[1];
    }

    // ── Durée — 4 stratégies ────────────────────────────────────────────────
    // Stratégie 1 : label explicite "Durée : Xh YY"
    m = htmlBody.match(/[Dd]ur[ée]{1,2}[^:]*:\s*(\d{1,2})\s*h\s*(\d{2})\s*(?:min)?/i)
     || plainBody.match(/[Dd]ur[ée]{1,2}[^:]*:\s*(\d{1,2})\s*h\s*(\d{2})\s*(?:min)?/i);
    if (m) {
      data.duree = `${parseInt(m[1])}h${m[2]}`;
    }

    // Stratégie 2 : "XhYY" standalone
    if (!data.duree) {
      const durPattern = /\b(\d{1,2})h(\d{2})\b/g;
      let candidate, match2;
      while ((match2 = durPattern.exec(htmlBody)) !== null) {
        const h     = parseInt(match2[1]);
        const min   = parseInt(match2[2]);
        const total = h * 60 + min;
        if (total >= 45 && total <= 270) {
          if (h < 7 || (h >= 1 && min > 0 && total <= 270)) {
            if (!candidate) candidate = match2;
          }
        }
      }
      if (candidate) data.duree = `${parseInt(candidate[1])}h${candidate[2]}`;
    }

    // Stratégie 3 : calcul depuis "Fin prévue à HH:MM"
    if (!data.duree && data.heure) {
      m = htmlBody.match(/Fin\s+pr[ée]vue\s+[àa]\s+(\d{1,2}:\d{2})/i)
       || plainBody.match(/Fin\s+pr[ée]vue\s+[àa]\s+(\d{1,2}:\d{2})/i);
      if (m) {
        const d1   = data.heure.split(':').map(Number);
        let d2     = m[1].split(':').map(Number);
        let min1   = d1[0] * 60 + d1[1];
        let min2   = d2[0] * 60 + d2[1];
        if (min2 < min1) min2 += 24 * 60;
        const dur  = Math.max(0, min2 - min1 - 15);
        if (dur >= 45) data.duree = `${Math.floor(dur/60)}h${String(dur%60).padStart(2,'0')}`;
      }
    }

    // Stratégie 4 : "XX minutes"
    if (!data.duree) {
      m = plainBody.match(/\b(\d{2,3})\s*min(?:utes?)?\b/i);
      if (m) {
        const total = parseInt(m[1]);
        if (total >= 45 && total <= 270) {
          data.duree = `${Math.floor(total/60)}h${String(total%60).padStart(2,'0')}`;
        }
      }
    }

    // ── Salle & siège ───────────────────────────────────────────────────────
    const roomPattern    = /Salle\s+([A-Z0-9][A-Z0-9 ]{0,18}?)(?:\s*[-–]\s*(?:Place|Siège)|<|$|\n)/i;
    const combinedPattern = /Salle\s+([A-Z0-9][A-Z0-9 ]{0,18?})\s*[-–]\s*(?:Place|Siège)\s+([A-Z]?\d+[A-Z]?)/i;

    let roomMatch = htmlBody.match(combinedPattern) || plainBody.match(combinedPattern);
    if (roomMatch) {
      data.salle = roomMatch[1].trim();
      data.siege = roomMatch[2].trim();
    } else {
      roomMatch = htmlBody.match(roomPattern) || plainBody.match(roomPattern);
      if (roomMatch) data.salle = roomMatch[1].trim();
      const seatMatch = htmlBody.match(/(?:Place|Siège)\s+([A-Z]?\d+[A-Z]?)/i)
                      || plainBody.match(/(?:Place|Siège)\s+([A-Z]?\d+[A-Z]?)/i);
      if (seatMatch) data.siege = seatMatch[1].trim();
    }

    // ── Langue ──────────────────────────────────────────────────────────────
    m = htmlBody.match(/>\s*(VF|VOST|VOSTFR|VO|VFQ)\s*</i);
    if (m) data.langue = m[1].toUpperCase().replace('VOSTFR', 'VOST');

    return data;
  } catch (e) {
    Logger.log('Erreur parseEmailBody : ' + e);
    return null;
  }
}

function getFilmsANoter() {
  try {
    const threads = GmailApp.search(
      'subject:"Confirmation de commande les cinémas Pathé" is:unread',
      0, 20
    );

    if (!threads || threads.length === 0) return [];

    const films = [];
    const processedKeys = new Set();

    for (const thread of threads) {
      const messages = thread.getMessages();
      for (const msg of messages) {
        try {
          if (!msg.isUnread()) continue;

          const email = parsePatheEmail(msg);
          if (email) {
            const uniqueKey = `${email.titre}|${email.date}|${email.heure}`;
            if (!processedKeys.has(uniqueKey)) {
              processedKeys.add(uniqueKey);

              const tmdb = getMovieDataFromTMDB(email.titre);

              const resolvedDuree = (email.duree && email.duree !== 'N/A')
                ? email.duree
                : (tmdb.duree && tmdb.duree !== 'N/A' ? tmdb.duree : null);

              films.push({
                ...email,
                ...tmdb,
                duree:     resolvedDuree,
                genre:     tmdb.genre || null,
                messageId: msg.getId()
              });

              if (films.length >= 10) break;
            }
          }
        } catch (e) {
          Logger.log('Erreur traitement message : ' + e);
        }
      }
      if (films.length >= 10) break;
    }

    films.sort((a, b) => {
      try {
        const dateA = new Date(a.date.split('/').reverse().join('-'));
        const dateB = new Date(b.date.split('/').reverse().join('-'));
        return dateA - dateB;
      } catch (e) { return 0; }
    });

    return films;
  } catch (e) {
    Logger.log('Erreur getFilmsANoter : ' + e);
    return [];
  }
}

// ==========================================
// TMDB
// ==========================================
function getMovieDataFromTMDB(titre) {
  try {
    let clean = decodeHtmlEntities(titre || '').trim();
    clean = clean.replace(/^["'«»]+|["'»«]+$/g, '').trim();
    clean = clean.replace(/[\u2018\u2019\u02BC]/g, "'");
    clean = clean.replace(/[^\p{L}\p{N}\s':\-.,!?&()]/gu, ' ').replace(/\s+/g, ' ').trim();

    if (!clean) return { affiche: null, duree: 'N/A', tmdbId: null, genre: null };

    const doSearch = (query) => {
      const url = 'https://api.themoviedb.org/3/search/movie'
        + '?api_key=' + TMDB_API_KEY
        + '&query='   + encodeURIComponent(query)
        + '&language=fr-FR';
      const res  = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const json = JSON.parse(res.getContentText());
      return (json && json.results) ? json.results : [];
    };

    const pickBest = (results) => {
      if (!results.length) return null;
      const norm = (s) => (s || '').toLowerCase().trim()
        .replace(/[\u2018\u2019\u02BC]/g, "'")
        .replace(/[^\p{L}\p{N}\s']/gu, '');
      const target = norm(clean);
      const scored = results.map(r => {
        const titleFR   = norm(r.title);
        const titleOrig = norm(r.original_title);
        let score = 0;
        if (titleFR === target || titleOrig === target)                          score = 100;
        else if (titleFR.includes(target)   || target.includes(titleFR))         score = 50;
        else if (titleOrig.includes(target) || target.includes(titleOrig))        score = 40;
        const year = r.release_date ? parseInt(r.release_date.substring(0, 4)) : 0;
        return { r, score, year };
      });
      scored.sort((a, b) => b.score - a.score || b.year - a.year);
      return scored[0].r;
    };

    let results = doSearch(clean);
    let movie   = pickBest(results);

    if (!movie) {
      const stripped = clean.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
      if (stripped !== clean) {
        results = doSearch(stripped);
        movie   = pickBest(results);
      }
    }

    if (!movie) return { affiche: null, duree: 'N/A', tmdbId: null, genre: null };

    const detailUrl = 'https://api.themoviedb.org/3/movie/' + movie.id
      + '?api_key=' + TMDB_API_KEY + '&language=fr-FR';
    const detail = UrlFetchApp.fetch(detailUrl, { muteHttpExceptions: true });
    const d = JSON.parse(detail.getContentText());

    const min    = d.runtime || 0;
    const duree  = min ? Math.floor(min/60) + 'h' + String(min%60).padStart(2,'0') : 'N/A';
    const poster = movie.poster_path
      ? 'https://image.tmdb.org/t/p/w500' + movie.poster_path
      : null;
    const primaryGenre = d.genres && d.genres.length > 0 ? d.genres[0].name : null;

    return { affiche: poster, duree: duree, tmdbId: movie.id, genre: primaryGenre };

  } catch (e) {
    Logger.log('Erreur TMDB : ' + e);
    return { affiche: null, duree: 'N/A', tmdbId: null, genre: null };
  }
}

function getMovieDetailsFromTMDB(titre) {
  try {
    const clean     = decodeHtmlEntities(titre).trim();
    const searchUrl = `https://api.themoviedb.org/3/search/movie?api_key=${TMDB_API_KEY}&query=${encodeURIComponent(clean)}&language=fr-FR`;
    const search    = UrlFetchApp.fetch(searchUrl, { muteHttpExceptions: true });
    const res       = JSON.parse(search.getContentText());

    if (!res.results || !res.results.length) return null;

    const movie   = res.results[0];
    const tmdbId  = movie.id;

    const detailUrl  = `https://api.themoviedb.org/3/movie/${tmdbId}?api_key=${TMDB_API_KEY}&language=fr-FR`;
    const detail     = UrlFetchApp.fetch(detailUrl, { muteHttpExceptions: true });
    const movieData  = JSON.parse(detail.getContentText());

    const creditsUrl  = `https://api.themoviedb.org/3/movie/${tmdbId}/credits?api_key=${TMDB_API_KEY}&language=fr-FR`;
    const credits     = UrlFetchApp.fetch(creditsUrl, { muteHttpExceptions: true });
    const creditsData = JSON.parse(credits.getContentText());

    const director = creditsData.crew.find(person => person.job === 'Director');
    const cast     = creditsData.cast.slice(0, 8).map(actor => ({
      name:         actor.name,
      character:    actor.character,
      profile_path: actor.profile_path
    }));

    let formattedReleaseDate = 'N/A';
    if (movieData.release_date) {
      try {
        const dateParts = movieData.release_date.split('-');
        if (dateParts.length === 3) {
          formattedReleaseDate = `${dateParts[2]}/${dateParts[1]}/${dateParts[0]}`;
        }
      } catch (e) { Logger.log('Erreur formatage date: ' + e); }
    }

    return {
      tmdbId:         tmdbId,
      title:          movieData.title,
      original_title: movieData.original_title,
      poster_path:    movieData.poster_path ? `https://image.tmdb.org/t/p/w500${movieData.poster_path}` : null,
      backdrop_path:  movieData.backdrop_path ? `https://image.tmdb.org/t/p/original${movieData.backdrop_path}` : null,
      release_date:   formattedReleaseDate,
      runtime:        movieData.runtime,
      director:       director ? director.name : 'Inconnu',
      cast:           cast,
      genres:         movieData.genres.map(g => g.name),
      overview:       movieData.overview
    };
  } catch (e) {
    Logger.log('Erreur getMovieDetailsFromTMDB : ' + e);
    return null;
  }
}

function getMoviePosterUrl(titre) {
  try {
    const tmdbData = getMovieDataFromTMDB(titre);
    if (tmdbData && tmdbData.affiche) return tmdbData.affiche;
    return null;
  } catch (e) {
    Logger.log('Erreur getMoviePosterUrl : ' + e);
    return null;
  }
}

function getMoviePosterUrls(titres) {
  const result = {};
  if (!titres || !titres.length) return result;

  const requests = titres.map(titre => {
    let clean = decodeHtmlEntities(titre || '').trim();
    clean = clean.replace(/^["'«»]+|["'»«]+$/g, '').trim();
    clean = clean.replace(/[\u2018\u2019\u02BC]/g, "'");
    clean = clean.replace(/[^\p{L}\p{N}\s':\-.,!?&()]/gu, ' ').replace(/\s+/g, ' ').trim();
    return {
      url: 'https://api.themoviedb.org/3/search/movie'
        + '?api_key=' + TMDB_API_KEY
        + '&query='   + encodeURIComponent(clean || titre)
        + '&language=fr-FR',
      muteHttpExceptions: true,
      _titre: titre,
      _clean: clean
    };
  });

  const CHUNK = 10;
  for (let i = 0; i < requests.length; i += CHUNK) {
    const chunk = requests.slice(i, i + CHUNK);
    let responses;
    try {
      responses = UrlFetchApp.fetchAll(chunk);
    } catch (e) {
      Logger.log('getMoviePosterUrls fetchAll error: ' + e);
      chunk.forEach(r => { result[r._titre] = null; });
      continue;
    }

    responses.forEach((resp, idx) => {
      const req = chunk[idx];
      try {
        const json  = JSON.parse(resp.getContentText());
        const items = (json && json.results) ? json.results : [];
        if (!items.length) { result[req._titre] = null; return; }

        const norm = s => (s || '').toLowerCase().trim()
          .replace(/[\u2018\u2019\u02BC]/g, "'")
          .replace(/[^\p{L}\p{N}\s']/gu, '');
        const target = norm(req._clean);

        const scored = items.map(r => {
          const tFR   = norm(r.title);
          const tOrig = norm(r.original_title);
          let score = 0;
          if (tFR === target || tOrig === target)                          score = 100;
          else if (tFR.includes(target)   || target.includes(tFR))         score = 50;
          else if (tOrig.includes(target) || target.includes(tOrig))        score = 40;
          const year = r.release_date ? parseInt(r.release_date.substring(0, 4)) : 0;
          return { r, score, year };
        });
        scored.sort((a, b) => b.score - a.score || b.year - a.year);

        const best = scored[0].r;
        result[req._titre] = best.poster_path
          ? 'https://image.tmdb.org/t/p/w342' + best.poster_path
          : null;
      } catch (e) {
        Logger.log('getMoviePosterUrls parse error for "' + req._titre + '": ' + e);
        result[req._titre] = null;
      }
    });
  }

  return result;
}

function getPosterByTitle(titre) {
  try {
    return getMoviePosterUrl(titre);
  } catch (e) {
    Logger.log('Erreur getPosterByTitle : ' + e);
    return null;
  }
}

function convertImageToBase64(url) {
  try {
    if (!url || typeof url !== 'string') return '';
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0' }
    });
    if (response.getResponseCode() !== 200) return '';
    const blob     = response.getBlob();
    const mimeType = blob.getContentType() || 'image/jpeg';
    const b64      = Utilities.base64Encode(blob.getBytes());
    return 'data:' + mimeType + ';base64,' + b64;
  } catch (e) {
    Logger.log('convertImageToBase64 error: ' + e);
    return '';
  }
}

// ==========================================
// NOTATION
// ==========================================
function getNextScreeningNumber() {
  try {
    const spreadsheetId = getUserSpreadsheetId();
    const sh    = SpreadsheetApp.openById(spreadsheetId);
    const sheet = sh.getSheetByName('DB');
    if (!sheet) return 1;
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return 1;
    const lastId = sheet.getRange(lastRow, 1).getValue();
    return (lastId || 0) + 1;
  } catch (e) {
    Logger.log('Erreur getNextScreeningNumber : ' + e);
    return 1;
  }
}

function saveNotation(d) {
  try {
    Logger.log('=== DEBUT saveNotation ===');
    Logger.log('d.titre = "' + d.titre + '"');

    if (!d.titre || !d.note) {
      return { success: false, error: 'Données manquantes : titre ou note' };
    }

    const spreadsheetId = getUserSpreadsheetId();
    const sh    = SpreadsheetApp.openById(spreadsheetId);
    const sheet = sh.getSheetByName('DB');
    if (!sheet) return { success: false, error: 'Feuille DB introuvable' };

    const lastRow = sheet.getLastRow();
    const newId   = lastRow > 1 ? sheet.getRange(lastRow, 1).getValue() + 1 : 1;

    // ── Récupérer affiche et tmdbId — avec fallback TMDB ─────────────────
    let affiche  = d.affiche  || null;
    let tmdbId   = d.tmdbId   || null;
    let tmdbGenre = null;

    if (!affiche || !tmdbId) {
      const tmdbData = getMovieDataFromTMDB(d.titre);
      if (tmdbData) {
        if (!affiche) affiche  = tmdbData.affiche;
        if (!tmdbId)  tmdbId   = tmdbData.tmdbId;
        tmdbGenre = tmdbData.genre;
      }
    }

    const rowData = [
      newId,
      d.titre        || 'Titre inconnu',
      d.date         || '',
      d.heure        || '',
      d.duree        || '',
      d.langue       || '',
      d.salle        || '',
      d.siege        || '',
      d.note         || '',
      d.coupDeCoeur  ? '1' : '0',
      d.genre        || tmdbGenre || '',   // ← correction : tmdbGenre au lieu de tmdbData?.genre
      d.extras       || '0',
      d.capucines    ? '1' : '0',
      d.comment      || '',
      affiche        || '',
      tmdbId         || ''
    ];

    sheet.appendRow(rowData);

    // Marquer l'email source comme lu
    if (d.messageId) {
      marquerEmailCommeLu(d.messageId);
    } else {
      Logger.log('⚠️ No messageId provided for "' + d.titre + '", email not marked as read');
    }

    return { success: true };
  } catch (e) {
    Logger.log('ERREUR saveNotation: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function marquerEmailCommeLu(messageId) {
  if (!messageId) {
    Logger.log('marquerEmailCommeLu: no messageId provided, skipping');
    return;
  }
  try {
    const msg = GmailApp.getMessageById(messageId);
    if (msg && msg.isUnread()) {
      msg.markRead();
      Logger.log('✅ Marked as read: message ID ' + messageId);
    }
  } catch (e) {
    Logger.log('ERREUR marquerEmailCommeLu: ' + e.toString());
  }
}

// ==========================================
// GENRES
// ==========================================
function getGenres(spreadsheetId) {
  try {
    const sheetId = spreadsheetId || getUserSpreadsheetId();
    const sh = SpreadsheetApp.openById(sheetId);
    const genresSheet = sh.getSheetByName('Genres');
    if (!genresSheet) return getDefaultGenres();
    const data   = genresSheet.getRange('A2:A').getValues();
    const genres = data.map(r => r[0]).filter(String);
    return genres.length > 0 ? genres : getDefaultGenres();
  } catch (e) {
    Logger.log('Erreur getGenres : ' + e);
    return getDefaultGenres();
  }
}

function getDefaultGenres() {
  return [
    'Action', 'Aventure', 'Comédie', 'Drame',
    'Fantastique', 'Horreur', 'Romance', 'Science-Fiction', 'Thriller'
  ];
}

// ==========================================
// HISTORIQUE
// ==========================================
function getFilmsGroupedByYear() {
  try {
    const sheetId = getUserSpreadsheetId();
    const sh      = SpreadsheetApp.openById(sheetId);
    const sheet   = sh.getSheetByName('DB');
    if (!sheet) return {};

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return {};

    const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
    const out  = {};

    data.forEach((r) => {
      if (!r[1]) return;
      const dateStr = r[2] ? r[2].toString() : '';
      const yMatch  = dateStr.match(/\d{4}/);
      const y       = yMatch ? yMatch[0] : 'Unknown';
      if (!out[y]) out[y] = [];
      out[y].push({
        titre:        decodeHtmlEntities(r[1]),
        note:         r[8],
        coupDeCoeur:  r[9] == 1,
        capucines:    r[12] === true || r[12] === 1 || r[12] === '1' || r[12] === 'TRUE',
        dateComplete: dateStr,
        genre:        r[10] ? decodeHtmlEntities(r[10].toString()) : '',
      });
    });

    for (const year in out) {
      out[year].sort((a, b) => {
        const dateA = new Date(a.dateComplete.split('/').reverse().join('-'));
        const dateB = new Date(b.dateComplete.split('/').reverse().join('-'));
        return dateB - dateA;
      });
    }

    return out;
  } catch (e) {
    Logger.log('Erreur dans getFilmsGroupedByYear: ' + e);
    return {};
  }
}

// ==========================================
// DASHBOARD
// ==========================================
function getDashboardStats(filterType, filterValue) {
  try {
    const sheetId = getUserSpreadsheetId();
    const sh      = SpreadsheetApp.openById(sheetId);
    const sheet   = sh.getSheetByName('DB');
    if (!sheet) return getEmptyStats();

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return getEmptyStats();

    const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
    let total = 0, sumNote = 0, fav = 0, cap = 0, sumDur = 0, extras = 0;
    const ratingD = {1:0, 2:0, 3:0, 4:0, 5:0};
    const genreD = {}, langD = {}, roomD = {}, seatD = {}, years = new Set(), months = new Set();
    const monthlyData = {};
    let prevPeriodTotal = 0, prevPeriodSumNote = 0, prevPeriodSumDur = 0, prevPeriodExtras = 0;
    let prevPeriodFav = 0, prevPeriodCap = 0;

    let compareYear = null, compareMonth = null;
    if (filterType === 'year' && filterValue) {
      compareYear = String(parseInt(filterValue) - 1);
    } else if (filterType === 'month' && filterValue) {
      const [year, month] = filterValue.split('-').map(Number);
      const prevDate = new Date(year, month - 2, 1);
      compareYear  = prevDate.getFullYear().toString();
      compareMonth = String(prevDate.getMonth() + 1).padStart(2, '0');
    }

    data.forEach(r => {
      if (!r[1]) return;
      const dateValue = r[2];
      if (!(dateValue instanceof Date)) return;

      const year     = dateValue.getFullYear().toString();
      const month    = String(dateValue.getMonth() + 1).padStart(2, '0');
      const monthKey = `${year}-${month}`;

      years.add(year);
      months.add(monthKey);
      if (!monthlyData[year]) monthlyData[year] = [0,0,0,0,0,0,0,0,0,0,0,0];

      let shouldInclude = false, shouldIncludeInPrevPeriod = false;
      if (filterType === 'all' || !filterType) {
        shouldInclude = true;
      } else if (filterType === 'year' && filterValue) {
        if (year === filterValue) shouldInclude = true;
        if (compareYear && year === compareYear) shouldIncludeInPrevPeriod = true;
      } else if (filterType === 'month' && filterValue) {
        if (monthKey === filterValue) shouldInclude = true;
        if (compareYear && compareMonth && monthKey === `${compareYear}-${compareMonth}`) shouldIncludeInPrevPeriod = true;
      }

      if (filterType !== 'month') monthlyData[year][dateValue.getMonth()]++;

      if (shouldIncludeInPrevPeriod) {
        prevPeriodTotal++;
        prevPeriodSumNote += Number(r[8]) || 0;
        prevPeriodFav     += Number(r[9]) || 0;
        prevPeriodCap     += Number(r[12]) || 0;
        const durStr2 = r[4] ? r[4].toString() : '';
        const durM2   = durStr2.match(/(\d+)h(\d+)/);
        if (durM2) prevPeriodSumDur += parseInt(durM2[1]) * 60 + parseInt(durM2[2]);
        prevPeriodExtras += Number(r[11]) || 0;
      }

      if (!shouldInclude) return;

      total++;
      sumNote += Number(r[8]) || 0;
      fav     += Number(r[9]) || 0;
      cap     += Number(r[12]) || 0;

      const durStr = r[4] ? r[4].toString() : '';
      const durMatch = durStr.match(/(\d+)h(\d+)/);
      if (durMatch) sumDur += parseInt(durMatch[1]) * 60 + parseInt(durMatch[2]);

      extras += Number(r[11]) || 0;

      const rating = Math.round(r[8]);
      if (rating >= 1 && rating <= 5) ratingD[rating] = (ratingD[rating] || 0) + 1;

      if (r[10]) genreD[r[10]] = (genreD[r[10]] || 0) + 1;
      if (r[5])  langD[r[5]]   = (langD[r[5]]   || 0) + 1;
      if (r[6])  roomD[r[6]]   = (roomD[r[6]]   || 0) + 1;
      if (r[7])  seatD[r[7]]   = (seatD[r[7]]   || 0) + 1;
    });

    const avgRating      = total             ? sumNote / total                     : 0;
    const avgDuration    = total             ? Math.round(sumDur / total)           : 0;
    const prevAvgRating  = prevPeriodTotal   ? prevPeriodSumNote / prevPeriodTotal  : 0;
    const prevAvgDuration= prevPeriodTotal   ? Math.round(prevPeriodSumDur / prevPeriodTotal) : 0;

    const totalHours  = getTotalCinemaHours(filterType, filterValue);
    const dailyRecord = getDailyRecord(filterType, filterValue);
    const streakData  = getStreakData();

    return {
      totalFilms:         total,
      averageRating:      avgRating,
      favoriteCount:      fav,
      capucinesCount:     cap,
      averageDuration:    avgDuration,
      totalExtrasSpent:   extras,
      ratingDistribution: ratingD,
      genreDistribution:  genreD,
      languageDistribution: langD,
      favRoom:            getMaxShare(roomD),
      favSeat:            getMaxShare(seatD),
      availableYears:     [...years].sort((a, b) => b - a),
      availableMonths:    [...months].sort((a, b) => b.localeCompare(a)),
      monthlyData:        monthlyData,
      yoyComparison: {
        films:     prevPeriodTotal > 0 ? total      - prevPeriodTotal   : null,
        rating:    prevPeriodTotal > 0 ? avgRating  - prevAvgRating     : null,
        duration:  prevPeriodTotal > 0 ? avgDuration- prevAvgDuration   : null,
        extras:    prevPeriodTotal > 0 ? extras     - prevPeriodExtras  : null,
        favoris:   prevPeriodTotal > 0 ? fav        - prevPeriodFav     : null,
        capucines: prevPeriodTotal > 0 ? cap        - prevPeriodCap     : null
      },
      totalCinemaHours: totalHours,
      dailyRecord:      dailyRecord,
      streakData:       streakData
    };
  } catch (e) {
    Logger.log('Erreur getDashboardStats : ' + e);
    return getEmptyStats();
  }
}

function getEmptyStats() {
  return {
    totalFilms: 0, averageRating: 0, favoriteCount: 0, capucinesCount: 0,
    averageDuration: 0, totalExtrasSpent: 0,
    ratingDistribution: {1:0,2:0,3:0,4:0,5:0},
    genreDistribution: {}, languageDistribution: {},
    favRoom: null, favSeat: null, availableYears: [], monthlyData: {},
    yoyComparison: { films:null, rating:null, duration:null, extras:null, favoris:null, capucines:null },
    totalCinemaHours: 0, dailyRecord: null,
    streakData: { currentStreak:0, longestStreak:0, longestStreakEnd:null }
  };
}

function getMaxShare(obj) {
  const tot = Object.values(obj).reduce((s, v) => s + v, 0);
  if (!tot) return null;
  let maxV = 0, maxK = '';
  Object.entries(obj).forEach(([k, v]) => { if (v > maxV) { maxV = v; maxK = k; } });
  return { name: maxK, share: Math.round((maxV / tot) * 100) };
}

// ==========================================
// UTILITAIRES
// ==========================================
function decodeHtmlEntities(text) {
  if (!text) return text;
  return text
    .replace(/&#x27;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/&lt;/g,   '<')
    .replace(/&gt;/g,   '>')
    .replace(/&amp;/g,  '&');
}

function formatDate(dateStr) {
  if (!dateStr) return 'XX/XX/XXXX';
  const p = dateStr.toString().split(/[\/\-]/);
  if (p.length === 3) {
    let d = p[0].padStart(2, '0');
    let m = p[1].padStart(2, '0');
    let y = p[2];
    if (y.length === 2) y = '20' + y;
    return `${d}/${m}/${y}`;
  }
  try {
    const d = new Date(dateStr);
    if (!isNaN(d.getTime())) {
      return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
    }
  } catch (e) { Logger.log('Erreur formatDate : ' + e); }
  return dateStr;
}

function getTotalCinemaHours(filterType, filterValue) {
  try {
    const sheetId = getUserSpreadsheetId();
    const sh      = SpreadsheetApp.openById(sheetId);
    const sheet   = sh.getSheetByName('DB');
    if (!sheet) return 0;
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return 0;

    const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
    let totalMinutes = 0;

    data.forEach(r => {
      if (!r[1]) return;
      const dateValue = r[2];
      if (!(dateValue instanceof Date)) return;
      const year     = dateValue.getFullYear().toString();
      const month    = String(dateValue.getMonth() + 1).padStart(2, '0');
      const monthKey = `${year}-${month}`;

      let shouldInclude = false;
      if (filterType === 'all' || !filterType)                         shouldInclude = true;
      else if (filterType === 'year'  && year     === filterValue)     shouldInclude = true;
      else if (filterType === 'month' && monthKey === filterValue)     shouldInclude = true;
      if (!shouldInclude) return;

      const durStr   = r[4] ? r[4].toString() : '';
      const durMatch = durStr.match(/(\d+)h(\d+)/);
      if (durMatch) totalMinutes += parseInt(durMatch[1]) * 60 + parseInt(durMatch[2]);
    });

    return Math.round(totalMinutes / 60 * 10) / 10;
  } catch (e) {
    Logger.log('Erreur getTotalCinemaHours : ' + e);
    return 0;
  }
}

function getDailyRecord(filterType, filterValue) {
  try {
    const sheetId = getUserSpreadsheetId();
    const sh      = SpreadsheetApp.openById(sheetId);
    const sheet   = sh.getSheetByName('DB');
    if (!sheet) return null;
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return null;

    const data     = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
    const dayCount = {}, dayDates = {};

    data.forEach(r => {
      if (!r[1]) return;
      const dateValue = r[2];
      if (!(dateValue instanceof Date)) return;
      const year     = dateValue.getFullYear().toString();
      const month    = String(dateValue.getMonth() + 1).padStart(2, '0');
      const monthKey = `${year}-${month}`;

      let shouldInclude = false;
      if (filterType === 'all' || !filterType)                         shouldInclude = true;
      else if (filterType === 'year'  && year     === filterValue)     shouldInclude = true;
      else if (filterType === 'month' && monthKey === filterValue)     shouldInclude = true;
      if (!shouldInclude) return;

      const day      = String(dateValue.getDate()).padStart(2, '0');
      const monthNum = String(dateValue.getMonth() + 1).padStart(2, '0');
      const dateKey  = `${day}/${monthNum}/${dateValue.getFullYear()}`;
      dayCount[dateKey] = (dayCount[dateKey] || 0) + 1;
      if (!dayDates[dateKey]) dayDates[dateKey] = dateValue;
    });

    if (Object.keys(dayCount).length === 0) return null;
    const maxCount = Math.max(...Object.values(dayCount));
    if (maxCount === 1) return null;

    const daysWithMax = Object.keys(dayCount).filter(d => dayCount[d] === maxCount);
    let mostRecentDate = null, mostRecentDateKey = null;
    daysWithMax.forEach(dateKey => {
      const cur = dayDates[dateKey];
      if (!mostRecentDate || cur > mostRecentDate) { mostRecentDate = cur; mostRecentDateKey = dateKey; }
    });

    return { count: maxCount, date: mostRecentDateKey };
  } catch (e) {
    Logger.log('Erreur getDailyRecord : ' + e);
    return null;
  }
}

function getStreakData() {
  try {
    const sheetId = getUserSpreadsheetId();
    const sh      = SpreadsheetApp.openById(sheetId);
    const sheet   = sh.getSheetByName('DB');
    if (!sheet) return { currentStreak:0, longestStreak:0, longestStreakEnd:null, nextDeadline:null };
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { currentStreak:0, longestStreak:0, longestStreakEnd:null, nextDeadline:null };

    const data    = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
    const weekSet = new Set(), weekDates = {};

    data.forEach(r => {
      if (!r[1]) return;
      const dateValue = r[2];
      if (!(dateValue instanceof Date)) return;
      const onejan     = new Date(dateValue.getFullYear(), 0, 1);
      const weekNumber = Math.ceil((((dateValue - onejan) / 86400000) + onejan.getDay() + 1) / 7);
      const weekKey    = `${dateValue.getFullYear()}-W${String(weekNumber).padStart(2,'0')}`;
      weekSet.add(weekKey);
      if (!weekDates[weekKey] || dateValue > weekDates[weekKey]) weekDates[weekKey] = dateValue;
    });

    if (weekSet.size === 0) return { currentStreak:0, longestStreak:0, longestStreakEnd:null, nextDeadline:null };

    const weeks  = Array.from(weekSet).sort();
    const today  = new Date();
    const onejan = new Date(today.getFullYear(), 0, 1);
    const curWN  = Math.ceil((((today - onejan) / 86400000) + onejan.getDay() + 1) / 7);
    const curWK  = `${today.getFullYear()}-W${String(curWN).padStart(2,'0')}`;

    let currentStreak = 0, checkWeek = curWK;
    if (!weekSet.has(checkWeek)) {
      const [y, w] = checkWeek.split('-W').map(Number);
      checkWeek = (w - 1 < 1) ? `${y-1}-W52` : `${y}-W${String(w-1).padStart(2,'0')}`;
    }
    while (weekSet.has(checkWeek)) {
      currentStreak++;
      const [year, week] = checkWeek.split('-W').map(Number);
      checkWeek = (week - 1 < 1) ? `${year-1}-W52` : `${year}-W${String(week-1).padStart(2,'0')}`;
    }

    let nextDeadline = null;
    if (currentStreak > 0) {
      const lastSessionDate = weekDates[Array.from(weekSet).sort().pop()];
      const endOfWeek = new Date(lastSessionDate);
      const day  = lastSessionDate.getDay();
      const diff = (day === 0 ? 0 : 7 - day);
      endOfWeek.setDate(lastSessionDate.getDate() + diff);
      const deadlineDate = new Date(endOfWeek);
      deadlineDate.setDate(endOfWeek.getDate() + 7);
      deadlineDate.setHours(23, 59, 59, 999);
      if (today <= deadlineDate) {
        nextDeadline = `${String(deadlineDate.getDate()).padStart(2,'0')}/${String(deadlineDate.getMonth()+1).padStart(2,'0')}/${deadlineDate.getFullYear()}`;
      }
    }

    let longestStreak = 1, longestStreakEnd = null, tempStreak = 1;
    for (let i = 1; i < weeks.length; i++) {
      const [prevYear, prevWeek] = weeks[i-1].split('-W').map(Number);
      const [currYear, currWeek] = weeks[i].split('-W').map(Number);
      const isConsecutive = (currYear === prevYear && currWeek === prevWeek + 1)
                         || (currYear === prevYear + 1 && currWeek === 1 && (prevWeek === 52 || prevWeek === 53));
      if (isConsecutive) {
        tempStreak++;
      } else {
        if (tempStreak >= longestStreak) { longestStreak = tempStreak; longestStreakEnd = weeks[i-1]; }
        tempStreak = 1;
      }
    }
    if (tempStreak >= longestStreak) { longestStreak = tempStreak; longestStreakEnd = weeks[weeks.length - 1]; }

    let formattedLongestEndDate = null;
    if (longestStreakEnd && weekDates[longestStreakEnd]) {
      const lEnd = weekDates[longestStreakEnd];
      formattedLongestEndDate = `${String(lEnd.getDate()).padStart(2,'0')}/${String(lEnd.getMonth()+1).padStart(2,'0')}/${lEnd.getFullYear()}`;
    }

    return {
      currentStreak:    currentStreak,
      longestStreak:    longestStreak,
      longestStreakEnd: formattedLongestEndDate,
      nextDeadline:     nextDeadline
    };
  } catch (e) {
    Logger.log('Erreur getStreakData : ' + e);
    return { currentStreak:0, longestStreak:0, longestStreakEnd:null, nextDeadline:null };
  }
}

// ==========================================
// REWIND MENSUEL
// ==========================================
function getMonthlyRewindData() {
  try {
    const sheetId = getUserSpreadsheetId();
    const sh      = SpreadsheetApp.openById(sheetId);
    const sheet   = sh.getSheetByName('DB');
    if (!sheet) return {};
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return {};

    const data        = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
    const monthlyData = {};

    data.forEach(r => {
      if (!r[1] || !(r[2] instanceof Date)) return;
      const date = r[2];
      const key  = `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'00')}`;
      if (!monthlyData[key]) monthlyData[key] = [];
      monthlyData[key].push({
        titre:       decodeHtmlEntities(r[1]),
        duree:       r[4]  || '',
        langue:      r[5]  || '',
        salle:       r[6]  ? r[6].toString().trim() : '',
        siege:       r[7]  ? r[7].toString().trim() : '',
        note:        Number(r[8]) || 0,
        coupDeCoeur: r[9]  == 1,
        genre:       r[10] || '',
        extras:      Number(r[11]) || 0,
        capucines:   (r[12] == 1 || r[12] === '1' || r[12] === true || r[12] === 'TRUE')
      });
    });

    const result = {};
    for (const monthKey in monthlyData) {
      const films = monthlyData[monthKey];
      if (!films.length) continue;

      let totalDuration = 0, totalNote = 0, favoriteCount = 0, totalExtras = 0;
      const languageDistribution = {}, genreDistribution = {};
      let bestMovie = null, worstMovie = null;

      films.forEach(film => {
        const m = film.duree.match(/(\d+)h(\d+)/);
        if (m) totalDuration += parseInt(m[1]) * 60 + parseInt(m[2]);
        totalNote   += film.note;
        totalExtras += film.extras;
        if (film.coupDeCoeur) favoriteCount++;
        if (film.langue) languageDistribution[film.langue] = (languageDistribution[film.langue] || 0) + 1;
        if (film.genre)  genreDistribution[film.genre]     = (genreDistribution[film.genre]     || 0) + 1;
        if (!bestMovie  || film.note > bestMovie.note)  bestMovie  = film;
        if (!worstMovie || film.note < worstMovie.note) worstMovie = film;
      });

      const seatCount = {};
      films.forEach(function(film) {
        const s = (film.siege || '').toString().trim();
        if (s && s.toLowerCase() !== 'libre' && s !== '-') {
          seatCount[s] = (seatCount[s] || 0) + 1;
        }
      });
      let favSeat = null;
      const seatEntries = Object.entries(seatCount);
      if (seatEntries.length > 0) {
        seatEntries.sort((a, b) => b[1] - a[1]);
        favSeat = { name: seatEntries[0][0], share: Math.round((seatEntries[0][1] / films.length) * 100) };
      }

      result[monthKey] = {
        films:                films,
        totalFilms:           films.length,
        averageRating:        totalNote / films.length,
        averageDuration:      Math.round(totalDuration / films.length),
        totalDuration:        totalDuration,
        favoriteCount:        favoriteCount,
        totalExtras:          totalExtras,
        languageDistribution: languageDistribution,
        genreDistribution:    genreDistribution,
        bestMovie:            bestMovie,
        worstMovie:           worstMovie,
        favSeat:              favSeat
      };
    }

    return result;
  } catch (e) {
    Logger.log('Rewind error: ' + e);
    return {};
  }
}

function getAvailableMonths() {
  try {
    const sheetId = getUserSpreadsheetId();
    const sh      = SpreadsheetApp.openById(sheetId);
    const sheet   = sh.getSheetByName('DB');
    if (!sheet) return {};
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return {};

    const data   = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
    const months = {};

    data.forEach(r => {
      if (!r[1] || !(r[2] instanceof Date)) return;
      const date     = r[2];
      const year     = date.getFullYear();
      const month    = date.getMonth() + 1;
      const monthKey = `${year}-${String(month).padStart(2,'0')}`;
      if (!months[monthKey]) {
        months[monthKey] = { year: year, month: month, label: getMonthName(month) + ' ' + year };
      }
    });

    return months;
  } catch (e) {
    Logger.log('Erreur getAvailableMonths: ' + e);
    return {};
  }
}

function getMonthName(month) {
  return ['Janvier','Février','Mars','Avril','Mai','Juin',
          'Juillet','Août','Septembre','Octobre','Novembre','Décembre'][month - 1];
}

// ==========================================
// FILMS RÉCENTS
// ==========================================
function getLastRatedMovie() {
  try {
    const sheetId = getUserSpreadsheetId();
    const ss      = SpreadsheetApp.openById(sheetId);
    const sheet   = ss.getSheetByName('DB');
    if (!sheet) return null;
    const values  = sheet.getDataRange().getValues();
    if (values.length < 2) return null;

    for (let i = values.length - 1; i >= 1; i--) {
      const row   = values[i];
      const titre = row[1];
      if (titre && titre.toString().trim() !== '') {
        return {
          timestamp:   row[0],
          titre:       titre.toString(),
          date:        row[2] ? (row[2] instanceof Date ? Utilities.formatDate(row[2],'GMT+1','dd/MM/yyyy') : row[2].toString()) : '-',
          duree:       row[4]  ? row[4].toString()  : '-',
          note:        parseFloat(row[8])  || 0,
          coupDeCoeur: (row[9] == 1 || row[9] === '1' || row[9] === 'OUI'),
          genre:       row[10] ? row[10].toString().trim() : 'Default',
          capucines:   (row[12] == 1 || row[12] === '1' || row[12] === true || row[12] === 'TRUE'),
          commentaire: row[13] ? row[13].toString() : ''
        };
      }
    }
    return null;
  } catch (e) {
    Logger.log('Erreur getLastRatedMovie: ' + e);
    return null;
  }
}

function getMovieByIndex(index) {
  try {
    const sheetId = getUserSpreadsheetId();
    const ss      = SpreadsheetApp.openById(sheetId);
    const sheet   = ss.getSheetByName('DB');
    if (!sheet) return null;
    const row = sheet.getRange(index + 1, 1, 1, 16).getValues()[0];
    if (!row[1]) return null;
    return {
      timestamp:   row[0],
      titre:       row[1].toString(),
      date:        row[2] ? (row[2] instanceof Date ? Utilities.formatDate(row[2],'GMT+1','dd/MM/yyyy') : row[2].toString()) : '-',
      duree:       row[4]  ? row[4].toString()  : '-',
      note:        parseFloat(row[8])  || 0,
      coupDeCoeur: (row[9] == 1 || row[9] === '1' || row[9] === 'OUI'),
      genre:       row[10] ? row[10].toString().trim() : 'Default',
      capucines:   (row[12] == 1 || row[12] === '1' || row[12] === true || row[12] === 'TRUE'),
      commentaire: row[13] ? row[13].toString() : ''
    };
  } catch (e) {
    Logger.log('Erreur getMovieByIndex: ' + e);
    return null;
  }
}

function getAllRatedMovies() {
  try {
    const sheetId = getUserSpreadsheetId();
    const ss      = SpreadsheetApp.openById(sheetId);
    const sheet   = ss.getSheetByName('DB');
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const values = sheet.getRange(1, 1, lastRow, 16).getValues();
    const movies = [];

    for (let i = values.length - 1; i >= 1; i--) {
      const row   = values[i];
      const titre = row[1];
      if (titre && titre.toString().trim() !== '') {
        movies.push({
          index:       i,
          timestamp:   row[0],
          titre:       titre.toString(),
          date:        row[2] ? (row[2] instanceof Date ? Utilities.formatDate(row[2],'GMT+1','dd/MM/yyyy') : row[2].toString()) : '-',
          duree:       row[4]  ? row[4].toString()  : '-',
          langue:      row[5]  ? row[5].toString()  : '-',
          salle:       row[6]  ? row[6].toString()  : '-',
          note:        parseFloat(row[8])  || 0,
          coupDeCoeur: (row[9] == 1 || row[9] === '1' || row[9] === 'OUI'),
          genre:       row[10] ? row[10].toString().trim() : 'Default',
          commentaire: row[13] ? row[13].toString() : '',
          capucines:   (row[12] == 1 || row[12] === '1' || row[12] === true || row[12] === 'TRUE'),
          affiche:     row[14] ? row[14].toString() : '',
          tmdbId:      row[15] ? row[15].toString() : ''
        });
      }
    }

    return movies;
  } catch (e) {
    Logger.log('Erreur getAllRatedMovies: ' + e);
    return [];
  }
}

function getDashboardPosters(n) {
  try {
    const sheetId = getUserSpreadsheetId();
    const sh      = SpreadsheetApp.openById(sheetId);
    const sheet   = sh.getSheetByName('DB');
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    const data    = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
    const results = [];

    for (let i = data.length - 1; i >= 0 && results.length < n; i--) {
      const r = data[i];
      if (!r[1]) continue;
      results.push({
        titre:       r[1]  ? r[1].toString() : '',
        date:        r[2]  ? (r[2] instanceof Date ? Utilities.formatDate(r[2],'GMT+1','dd/MM/yyyy') : r[2].toString()) : '',
        note:        parseFloat(r[8])  || 0,
        coupDeCoeur: (r[9] == 1 || r[9] === '1'),
        genre:       r[10] ? r[10].toString() : '',
        affiche:     r[14] ? r[14].toString() : '',
        tmdbId:      r[15] ? r[15].toString() : ''
      });
    }

    return results;
  } catch (e) {
    Logger.log('getDashboardPosters error: ' + e);
    return [];
  }
}

// ==========================================
// PROFIL CINÉPHILE
// ==========================================
function getMovieProfile(filterType, filterValue) {
  try {
    const sheetId = getUserSpreadsheetId();
    const sh      = SpreadsheetApp.openById(sheetId);
    const sheet   = sh.getSheetByName('DB');
    if (!sheet) return getEmptyProfile();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return getEmptyProfile();

    const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
    const genreScores = {}, languagePrefs = {}, durationRatings = [];
    let recommendedCount = 0, totalFavorites = 0, totalFilms = 0;

    data.forEach(row => {
      if (!row[1]) return;
      const dateValue = row[2];
      if (!(dateValue instanceof Date)) return;
      const year     = dateValue.getFullYear().toString();
      const month    = String(dateValue.getMonth() + 1).padStart(2, '0');
      const monthKey = `${year}-${month}`;

      let shouldInclude = false;
      if (filterType === 'all' || !filterType)                         shouldInclude = true;
      else if (filterType === 'year'  && year     === filterValue)     shouldInclude = true;
      else if (filterType === 'month' && monthKey === filterValue)     shouldInclude = true;
      if (!shouldInclude) return;

      totalFilms++;
      const genre      = row[10];
      const note       = Number(row[8]) || 0;
      const isFavorite = row[9] == 1;
      const language   = row[5];

      if (note >= 3) recommendedCount++;
      if (isFavorite) totalFavorites++;

      if (genre) {
        if (!genreScores[genre]) genreScores[genre] = { total:0, count:0, favorites:0 };
        genreScores[genre].total += note;
        genreScores[genre].count++;
        if (isFavorite) genreScores[genre].favorites++;
      }

      if (language) languagePrefs[language] = (languagePrefs[language] || 0) + 1;

      const durStr   = row[4] ? row[4].toString() : '';
      const durMatch = durStr.match(/(\d+)h(\d+)/);
      if (durMatch) durationRatings.push({ duration: parseInt(durMatch[1])*60 + parseInt(durMatch[2]), rating: note });
    });

    if (totalFilms === 0) return getEmptyProfile();

    const genreAverages = Object.entries(genreScores)
      .map(([genre, d]) => ({
        genre,
        avgRating:    d.total / d.count,
        count:        d.count,
        favoriteRate: (d.favorites / d.count) * 100
      }))
      .sort((a, b) => b.avgRating - a.avgRating);

    const topGenres = genreAverages.slice(0, 2);
    const preferredLanguage = Object.entries(languagePrefs).sort((a,b) => b[1] - a[1])[0];
    const highRated = durationRatings.filter(d => d.rating >= 4);
    const avgDurHigh = highRated.length > 0 ? highRated.reduce((acc, d) => acc + d.duration, 0) / highRated.length : 0;
    const durationPreference = getDurationCategory(avgDurHigh);
    const satisfactionRate   = ((recommendedCount / totalFilms) * 100).toFixed(1);
    const favoriteRate       = ((totalFavorites   / totalFilms) * 100).toFixed(1);

    const insights = generateMovieInsights(genreAverages, totalFilms, totalFavorites, satisfactionRate, durationPreference, preferredLanguage);

    return {
      topGenres:         topGenres,
      preferredLanguage: preferredLanguage ? preferredLanguage[0] : 'N/A',
      languageShare:     preferredLanguage ? Math.round((preferredLanguage[1] / totalFilms) * 100) : 0,
      durationPreference: durationPreference,
      favoriteRate:      parseFloat(favoriteRate),
      satisfactionRate:  parseFloat(satisfactionRate),
      insights:          insights,
      profileType:       getProfileType(genreAverages, totalFilms),
      totalFilms:        totalFilms
    };
  } catch (e) {
    Logger.log('Erreur getMovieProfile : ' + e);
    return getEmptyProfile();
  }
}

function getEmptyProfile() {
  return {
    topGenres: [], preferredLanguage: 'N/A', languageShare: 0,
    durationPreference: 'N/A', favoriteRate: 0, satisfactionRate: 0,
    insights: ['Commence à noter des films pour découvrir ton profil !'],
    profileType: 'Explorateur amateur', totalFilms: 0
  };
}

function getDurationCategory(avgMinutes) {
  if (!avgMinutes || avgMinutes === 0) return 'N/A';
  if (avgMinutes < 90)  return '< 1h30';
  if (avgMinutes < 120) return '1h30-2h';
  if (avgMinutes < 150) return '2h-2h30';
  return '> 2h30';
}

function getProfileType(genreAverages, totalFilms) {
  if (!genreAverages || genreAverages.length === 0) return 'Explorateur amateur';
  const topGenre = genreAverages[0].genre.toLowerCase();
  let genreType  = 'Cinéphile';
  if (topGenre.includes('drame'))           genreType = 'Dramaturge';
  else if (topGenre.includes('comédie'))    genreType = 'Amuseur';
  else if (topGenre.includes('thriller'))   genreType = 'Enquêteur';
  else if (topGenre.includes('action'))     genreType = 'Aventurier';
  else if (topGenre.includes('horreur'))    genreType = 'Frissonneur';
  else if (topGenre.includes('science'))    genreType = 'Visionnaire';
  else if (topGenre.includes('fantastique'))genreType = 'Rêveur';
  else if (topGenre.includes('romance'))    genreType = 'Romantique';
  else if (topGenre.includes('aventure'))   genreType = 'Explorateur';
  let level = 'amateur';
  if (totalFilms >= 200) level = 'expert';
  else if (totalFilms >= 100) level = 'chevronné';
  else if (totalFilms >= 50)  level = 'averti';
  return `${genreType} ${level}`;
}

function generateMovieInsights(genreAverages, totalFilms, favorites, satisfactionRate, durationPref, preferredLang) {
  const insights = [];
  if (genreAverages.length > 0) {
    const top = genreAverages[0];
    insights.push(`${getGenreEmoji(top.genre)} Passion : ${top.genre} (${top.avgRating.toFixed(1)}/5)`);
  }
  if (preferredLang && preferredLang[0] !== 'N/A') {
    const langEmoji = preferredLang[0] === 'VOST' ? '🌍' : '🇫🇷';
    insights.push(`${langEmoji} ${preferredLang[0]} à ${Math.round((preferredLang[1]/totalFilms)*100)}%`);
  }
  if (durationPref !== 'N/A') insights.push(`⏱️ Préférence : films ${durationPref.toLowerCase()}`);
  const favRate = (favorites / totalFilms) * 100;
  if (favRate > 30) insights.push(`❤️ Généreux : ${favRate.toFixed(0)}% de coups de cœur`);
  else if (favRate < 10) insights.push(`🎯 Sélectif : ${favRate.toFixed(0)}% de coups de cœur`);
  return insights.slice(0, 4);
}

function getGenreEmoji(genre) {
  const emojis = {
    'Action':'💥','Aventure':'🗺️','Comédie':'😂','Drame':'🎭',
    'Fantastique':'✨','Horreur':'👻','Romance':'💕','Science-Fiction':'🚀','Thriller':'🔪'
  };
  return emojis[genre] || '🎬';
}

// ==========================================
// INSCRIPTION & PROFIL UTILISATEUR
// ==========================================

function createUserProfile(formData) {
  try {
    Logger.log('=== START createUserProfile ===');

    const userEmail = Session.getActiveUser().getEmail();
    Logger.log('User email: ' + userEmail);

    if (formData.confirmEmail.toLowerCase() !== userEmail.toLowerCase()) {
      return { success: false, error: 'Email ne correspond pas à ton compte Google' };
    }

    let spreadsheetId = formData.spreadsheetId?.trim();

    if (!spreadsheetId) {
      const result = createUserSpreadsheet(formData.firstName);
      if (!result.success) return { success: false, error: 'Erreur création spreadsheet: ' + result.error };
      spreadsheetId = result.spreadsheetId;
    } else {
      const validation = validateSpreadsheet(spreadsheetId);
      if (!validation.success) return validation;
    }

    const ss = SpreadsheetApp.openById(spreadsheetId);
    setupUserSheets(ss);

    let avatarUrl = formData.avatar || AVATAR_PRESETS[0];
    if (!AVATAR_PRESETS.includes(avatarUrl)) avatarUrl = AVATAR_PRESETS[0];

    const preferences = {
      theme:         formData.preferences?.theme         || 'dark-grey',
      ratingScale:   formData.preferences?.ratingScale   || '5',
      halfPoints:    formData.preferences?.halfPoints    === true,
      storyModal:    formData.preferences?.storyModal    === true,
      shareFunction: formData.preferences?.shareFunction === true,
      rewind:        formData.preferences?.rewind        === true,
      capucines:     formData.preferences?.capucines     === true,
      rating:    true,
      history:   true,
      dashboard: true
    };

    const profileData = {
      email:         userEmail,
      firstName:     formData.firstName,
      spreadsheetId: spreadsheetId,
      avatar:        avatarUrl,
      preferences:   preferences,
      creationDate:  new Date().toISOString(),
      lastModified:  new Date().toISOString()
    };

    PropertiesService.getUserProperties().setProperty('userProfile', JSON.stringify(profileData));

    let trackingSaved = false, lastError = null;
    for (let attempt = 1; attempt <= 5; attempt++) {
      try {
        const trackingResult = saveToTrackingSheet(profileData);
        if (trackingResult.success) { trackingSaved = true; break; }
        else { lastError = trackingResult.error; }
      } catch (e) { lastError = e.toString(); }
      if (attempt < 5) Utilities.sleep(attempt * 500);
    }

    if (!trackingSaved) {
      PropertiesService.getUserProperties().deleteAllProperties();
      return { success: false, error: 'Impossible d\'enregistrer dans le système de suivi. Erreur: ' + lastError };
    }

    const verifyProfile = getProfileFromTrackingSheet(userEmail);
    if (!verifyProfile) {
      PropertiesService.getUserProperties().deleteAllProperties();
      return { success: false, error: 'Compte créé mais non trouvé dans le système de suivi.' };
    }

    Logger.log('=== Registration SUCCESS ===');
    return { success: true, appUrl: ScriptApp.getService().getUrl(), profile: profileData };

  } catch (e) {
    Logger.log('❌ createUserProfile CRITICAL ERROR: ' + e.toString());
    return { success: false, error: 'Erreur système: ' + e.message };
  }
}

function saveToTrackingSheet(profileData) {
  try {
    const trackingSheet = SpreadsheetApp.openById(USER_TRACKING_SHEET_ID);
    let sheet = trackingSheet.getSheetByName('Users');

    if (!sheet) {
      sheet = trackingSheet.insertSheet('Users');
      const headers = ['First Name','Email','Spreadsheet ID','Avatar','Theme',
                       'Rating Scale','Half Points','Story Modal','Share Function',
                       'Rewind','Capucines','Creation Date','Last Modified'];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }

    const rowData = [
      profileData.firstName,
      profileData.email,
      profileData.spreadsheetId,
      profileData.avatar || AVATAR_PRESETS[0],
      profileData.preferences.theme         || 'dark-grey',
      profileData.preferences.ratingScale   || '5',
      profileData.preferences.halfPoints    ? 'Yes' : 'No',
      profileData.preferences.storyModal    ? 'Yes' : 'No',
      profileData.preferences.shareFunction ? 'Yes' : 'No',
      profileData.preferences.rewind        ? 'Yes' : 'No',
      profileData.preferences.capucines     ? 'Yes' : 'No',
      profileData.creationDate || new Date().toISOString(),
      profileData.lastModified || new Date().toISOString()
    ];

    const data = sheet.getDataRange().getValues();
    let existingRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === profileData.email) { existingRow = i + 1; break; }
    }

    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    Logger.log('saveToTrackingSheet ERROR: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function getProfileFromTrackingSheet(userEmail) {
  try {
    const trackingSheet = SpreadsheetApp.openById(USER_TRACKING_SHEET_ID);
    const sheet = trackingSheet.getSheetByName('Users');
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toLowerCase() === userEmail.toLowerCase()) {
        const row = data[i];
        return {
          firstName:     row[0],
          email:         row[1],
          spreadsheetId: row[2],
          avatar:        row[3] || AVATAR_PRESETS[0],
          preferences: {
            theme:         row[4]  || 'dark-grey',
            ratingScale:   row[5]  || '5',
            halfPoints:    row[6]  === 'Yes',
            storyModal:    row[7]  === 'Yes',
            shareFunction: row[8]  === 'Yes',
            rewind:        row[9]  === 'Yes',
            capucines:     row[10] === 'Yes',
            rating:    true,
            history:   true,
            dashboard: true
          },
          creationDate: row[11],
          lastModified: row[12]
        };
      }
    }
    return null;
  } catch (e) {
    Logger.log('getProfileFromTrackingSheet ERROR: ' + e.toString());
    return null;
  }
}

function getUserProfile() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) return null;

    const cache    = CacheService.getUserCache();
    const cacheKey = 'user_profile_' + userEmail;
    const cached   = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);

    // Priorité 1 : UserProperties
    const userProps   = PropertiesService.getUserProperties();
    const propsProfile = userProps.getProperty('userProfile');
    if (propsProfile) {
      const profile = JSON.parse(propsProfile);
      cache.put(cacheKey, JSON.stringify(profile), 600);
      try { saveToTrackingSheet(profile); } catch (e) { Logger.log('Background sync failed: ' + e); }
      return profile;
    }

    // Priorité 2 : Tracking sheet
    const trackingProfile = getProfileFromTrackingSheet(userEmail);
    if (trackingProfile) {
      userProps.setProperty('userProfile', JSON.stringify(trackingProfile));
      cache.put(cacheKey, JSON.stringify(trackingProfile), 600);
      return trackingProfile;
    }

    return null;
  } catch (e) {
    Logger.log('getUserProfile ERROR: ' + e.toString());
    return null;
  }
}

function createUserSpreadsheet(firstName) {
  try {
    const ss            = SpreadsheetApp.create(`Cinema Tracker - ${firstName}`);
    const spreadsheetId = ss.getId();
    const dbSheet       = ss.getActiveSheet();
    dbSheet.setName('DB');

    const headers = ['ID','Titre','Date','Heure','Durée','Langue','Salle','Siège',
                     'Note','Coup de Coeur','Genre','Extras','Capucines','Commentaire','Affiche','TMDB ID'];
    dbSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

    const genresSheet    = ss.insertSheet('Genres');
    const defaultGenres  = [['Genre'],['Action'],['Aventure'],['Comédie'],['Drame'],
                             ['Fantastique'],['Horreur'],['Romance'],['Science-Fiction'],['Thriller']];
    genresSheet.getRange(1, 1, defaultGenres.length, 1).setValues(defaultGenres);
    genresSheet.getRange(1, 1).setFontWeight('bold');

    return { success: true, spreadsheetId: spreadsheetId, spreadsheetUrl: ss.getUrl() };
  } catch (e) {
    Logger.log('createUserSpreadsheet ERROR: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function setupUserSheets(ss) {
  try {
    let dbSheet = ss.getSheetByName('DB');
    if (!dbSheet) {
      dbSheet = ss.insertSheet('DB');
      const headers = ['ID','Titre','Date','Heure','Durée','Langue','Salle','Siège',
                       'Note','Coup de Coeur','Genre','Extras','Capucines','Commentaire','Affiche','TMDB ID'];
      dbSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    }
    let genresSheet = ss.getSheetByName('Genres');
    if (!genresSheet) {
      genresSheet = ss.insertSheet('Genres');
      const defaultGenres = [['Genre'],['Action'],['Aventure'],['Comédie'],['Drame'],
                              ['Fantastique'],['Horreur'],['Romance'],['Science-Fiction'],['Thriller']];
      genresSheet.getRange(1, 1, defaultGenres.length, 1).setValues(defaultGenres);
      genresSheet.getRange(1, 1).setFontWeight('bold');
    }
    return true;
  } catch (e) {
    Logger.log('setupUserSheets ERROR: ' + e.toString());
    return false;
  }
}

function validateSpreadsheet(spreadsheetId) {
  try {
    const ss      = SpreadsheetApp.openById(spreadsheetId);
    const dbSheet = ss.getSheetByName('DB');
    if (!dbSheet) return { success: false, error: "La feuille 'DB' est introuvable" };

    const headers         = dbSheet.getRange(1, 1, 1, 16).getValues()[0];
    const requiredHeaders = ['ID','Titre','Note','Genre'];
    const hasRequired     = requiredHeaders.every(h => headers.some(header => header.toString().includes(h)));
    if (!hasRequired) return { success: false, error: 'Structure des colonnes incorrecte' };

    return { success: true };
  } catch (e) {
    Logger.log('validateSpreadsheet ERROR: ' + e.toString());
    return { success: false, error: "Impossible d'accéder au spreadsheet: " + e.message };
  }
}

function updateUserPreferences(preferences) {
  try {
    const profile = getUserProfile();
    if (!profile) return { success: false, error: 'No profile found' };

    const fields = ['storyModal','shareFunction','rewind','theme','ratingScale','halfPoints','capucines'];
    fields.forEach(f => { if (preferences[f] !== undefined) profile.preferences[f] = preferences[f]; });
    profile.lastModified = new Date().toISOString();

    saveToTrackingSheet(profile);
    PropertiesService.getUserProperties().setProperty('userProfile', JSON.stringify(profile));
    CacheService.getUserCache().remove('user_profile_' + Session.getActiveUser().getEmail());

    return { success: true, profile: profile };
  } catch (e) {
    Logger.log('updateUserPreferences ERROR: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function updateUserFirstName(newFirstName) {
  try {
    const profile = getUserProfile();
    if (!profile) return { success: false, error: 'No profile found' };

    profile.firstName    = newFirstName;
    profile.lastModified = new Date().toISOString();

    saveToTrackingSheet(profile);
    PropertiesService.getUserProperties().setProperty('userProfile', JSON.stringify(profile));
    CacheService.getUserCache().remove('user_profile_' + Session.getActiveUser().getEmail());

    return { success: true, profile: profile };
  } catch (e) {
    Logger.log('updateUserFirstName ERROR: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function updateUserAvatar(avatarUrl) {
  try {
    if (!AVATAR_PRESETS.includes(avatarUrl)) return { success: false, error: 'Invalid avatar URL' };

    const profile = getUserProfile();
    if (!profile) return { success: false, error: 'No profile found' };

    profile.avatar       = avatarUrl;
    profile.lastModified = new Date().toISOString();

    saveToTrackingSheet(profile);
    PropertiesService.getUserProperties().setProperty('userProfile', JSON.stringify(profile));
    CacheService.getUserCache().remove('user_profile_' + Session.getActiveUser().getEmail());

    return { success: true, avatar: avatarUrl };
  } catch (e) {
    Logger.log('updateUserAvatar ERROR: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function getUserAvatar()         { return getUserProfile()?.avatar        || AVATAR_PRESETS[0]; }
function getUserSpreadsheetId()  { return getUserProfile()?.spreadsheetId || SPREADSHEET_ID; }
function isUserRegistered()      { return getUserProfile() !== null; }
function getUserFirstName()      { return getUserProfile()?.firstName      || 'User'; }

function getUserPreferences() {
  const profile = getUserProfile();
  if (profile?.preferences) {
    return {
      theme:         profile.preferences.theme         || 'dark-grey',
      ratingScale:   profile.preferences.ratingScale   || '5',
      halfPoints:    profile.preferences.halfPoints    === true,
      storyModal:    profile.preferences.storyModal    === true,
      shareFunction: profile.preferences.shareFunction === true,
      rewind:        profile.preferences.rewind        === true,
      capucines:     profile.preferences.capucines     === true,
      rating:    true,
      history:   true,
      dashboard: true
    };
  }
  return { theme:'dark-grey', ratingScale:'5', halfPoints:false, storyModal:false, shareFunction:false, rewind:false, capucines:false, rating:true, history:true, dashboard:true };
}

function deleteUserAccount() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) return { success: false, error: 'No user session' };

    const trackingSheet = SpreadsheetApp.openById(USER_TRACKING_SHEET_ID);
    const sheet = trackingSheet.getSheetByName('Users');
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][1] === userEmail) { sheet.deleteRow(i + 1); break; }
      }
    }

    PropertiesService.getUserProperties().deleteAllProperties();
    CacheService.getUserCache().remove('user_profile_' + userEmail);

    return { success: true, message: 'Account deleted successfully' };
  } catch (e) {
    Logger.log('deleteUserAccount ERROR: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ── Synchronisation depuis le front ────────────────────────────────────────
// Appelée par syncUserToTracking() dans le JS du front
function syncUserToTrackingSheet() {
  try {
    const profile = getUserProfile();
    if (!profile) return { success: false, error: 'No profile found' };
    const result = saveToTrackingSheet(profile);
    if (result.success) {
      return { success: true, message: 'Compte synchronisé avec succès !' };
    }
    return { success: false, error: result.error };
  } catch (e) {
    Logger.log('syncUserToTrackingSheet ERROR: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function isCapucinesFilm(titre) {
  try {
    const spreadsheetId = getUserSpreadsheetId();
    const ss    = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName('DB');
    if (!sheet) return false;
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return false;
    const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
    return data.some(r => r[1] && r[1].toString().trim() === titre.trim() && (r[12] == 1 || r[12] === '1'));
  } catch (e) {
    Logger.log('isCapucinesFilm error: ' + e);
    return false;
  }
}

function getNextUnratedFilm() {
  try {
    const films = getFilmsANoter();
    if (!films || films.length === 0) return null;
    const f = films[0];
    return { titre: f.titre||'', date: f.date||'', affiche: f.affiche||'', genre: f.genre||'', duree: f.duree||'' };
  } catch (e) {
    Logger.log('getNextUnratedFilm error: ' + e);
    return null;
  }
}

// ==========================================
// FESTIVALS
// ==========================================
const FESTIVALS_SPREADSHEET_ID = '1IAJU3Uum6e36WMTp0srXBq_o5hrkwZYuEF8htDQmlOA';

function getFestivalData() {
  try {
    const ss         = SpreadsheetApp.openById(FESTIVALS_SPREADSHEET_ID);
    const ceremonies = getCeremonies(ss);
    const nominees   = getNominees(ss);

    const grouped = {};
    nominees.forEach(n => {
      if (!grouped[n.ceremony_id]) grouped[n.ceremony_id] = {};
      if (!grouped[n.ceremony_id][n.category]) grouped[n.ceremony_id][n.category] = [];
      grouped[n.ceremony_id][n.category].push(n);
    });
    Object.keys(grouped).forEach(cid => {
      Object.keys(grouped[cid]).forEach(cat => {
        grouped[cid][cat].sort((a, b) => a.display_order - b.display_order);
      });
    });

    const result = ceremonies.map(c => ({ ...c, categories: grouped[c.ceremony_id] || {} }));
    return { success: true, data: result };
  } catch (e) {
    Logger.log('getFestivalData ERROR: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function getCeremonies(ss) {
  const sheet = ss.getSheetByName('Ceremonies');
  if (!sheet) throw new Error("Onglet 'Ceremonies' introuvable");
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1).filter(r => r[0]).map(r => ({
    ceremony_id:  r[0].toString().trim(),
    name:         r[1].toString().trim(),
    year:         r[2].toString().trim(),
    country:      r[3].toString().trim(),
    color_accent: r[4].toString().trim() || '#E8B200'
  }));
}

function getNominees(ss) {
  const sheet = ss.getSheetByName('Nominees');
  if (!sheet) throw new Error("Onglet 'Nominees' introuvable");
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1).filter(r => r[0]).map(r => ({
    ceremony_id:  r[0].toString().trim(),
    category:     r[1].toString().trim(),
    nominee_name: r[2].toString().trim(),
    film_title:   r[3].toString().trim(),
    media_type:   r[4].toString().trim().toLowerCase() || 'movie',
    display_order: parseInt(r[5]) || 0,
    image_url:    null
  }));
}

function fetchNomineeImage(nomineeName, filmTitle, mediaType, ceremonyYear) {
  try {
    if (mediaType === 'person') return fetchPersonImage(nomineeName, filmTitle, ceremonyYear);
    return fetchMovieImage(filmTitle || nomineeName, ceremonyYear);
  } catch (e) {
    Logger.log('fetchNomineeImage ERROR for "' + nomineeName + '": ' + e.toString());
    return null;
  }
}

function fetchPersonImage(name, filmTitle, ceremonyYear) {
  const url  = 'https://api.themoviedb.org/3/search/person?api_key=' + TMDB_API_KEY + '&query=' + encodeURIComponent(name) + '&language=fr-FR';
  const res  = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const json = JSON.parse(res.getContentText());
  if (json.results && json.results.length) {
    const person = json.results.sort((a, b) => b.popularity - a.popularity)[0];
    if (person.profile_path) return 'https://image.tmdb.org/t/p/w342' + person.profile_path;
  }
  if (filmTitle) return fetchMovieImage(filmTitle, ceremonyYear);
  return null;
}

function fetchMovieImage(title, ceremonyYear) {
  const url  = 'https://api.themoviedb.org/3/search/movie?api_key=' + TMDB_API_KEY + '&query=' + encodeURIComponent(title) + '&language=fr-FR';
  const res  = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const json = JSON.parse(res.getContentText());
  if (!json.results || !json.results.length) return null;

  const norm = s => (s||'').toLowerCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9\s]/g,'').replace(/\s+/g,' ');
  const target = norm(title);
  const exactMatches = json.results.filter(r => norm(r.title) === target || norm(r.original_title) === target);
  const pool = exactMatches.length > 0 ? exactMatches : [];
  if (!pool.length) return null;

  if (ceremonyYear) {
    const yearNum     = parseInt(ceremonyYear);
    const yearFiltered = pool.filter(r => { if (!r.release_date) return false; const y = parseInt(r.release_date.substring(0,4)); return y >= yearNum-1 && y <= yearNum; });
    if (yearFiltered.length > 0) {
      yearFiltered.sort((a, b) => new Date(b.release_date) - new Date(a.release_date));
      return 'https://image.tmdb.org/t/p/w342' + yearFiltered[0].poster_path;
    }
  }

  pool.sort((a, b) => {
    const da = a.release_date ? new Date(a.release_date) : new Date(0);
    const db = b.release_date ? new Date(b.release_date) : new Date(0);
    return db - da;
  });
  return pool[0].poster_path ? 'https://image.tmdb.org/t/p/w342' + pool[0].poster_path : null;
}

function fetchCeremonyImages(nominees) {
  if (!nominees || !nominees.length) return [];
  const CHUNK_SIZE = 10;
  const results    = [...nominees];

  for (let i = 0; i < nominees.length; i += CHUNK_SIZE) {
    const chunk = nominees.slice(i, i + CHUNK_SIZE);
    const requests = chunk.map(n => {
      const isPerson = n.media_type === 'person';
      const query    = isPerson ? n.nominee_name : (n.film_title || n.nominee_name);
      const endpoint = isPerson ? 'search/person' : 'search/movie';
      return {
        url: `https://api.themoviedb.org/3/${endpoint}?api_key=${TMDB_API_KEY}&query=${encodeURIComponent(query)}&language=fr-FR`,
        muteHttpExceptions: true,
        _idx: i + chunk.indexOf(n)
      };
    });

    let responses;
    try { responses = UrlFetchApp.fetchAll(requests); }
    catch (e) { Logger.log('fetchAll error: ' + e); continue; }

    responses.forEach((resp, ri) => {
      try {
        const json  = JSON.parse(resp.getContentText());
        const items = json.results || [];
        if (!items.length) return;
        const best    = items.sort((a, b) => b.popularity - a.popularity)[0];
        const isMovie = chunk[ri].media_type !== 'person';
        const path    = isMovie ? best.poster_path : best.profile_path;
        results[requests[ri]._idx].image_url = path ? 'https://image.tmdb.org/t/p/w342' + path : null;
      } catch (e) { Logger.log('Parse error: ' + e); }
    });
  }

  return results;
}

// ==========================================
// DEBUGGING
// ==========================================
function checkUserProfile() {
  const userEmail = Session.getActiveUser().getEmail();
  Logger.log('=== CHECKING PROFILE FOR: ' + userEmail + ' ===');
  const profile         = getUserProfile();
  const trackingProfile = getProfileFromTrackingSheet(userEmail);
  Logger.log('UserProperties profile: ' + (profile ? 'FOUND' : 'NOT FOUND'));
  Logger.log('Tracking sheet profile: ' + (trackingProfile ? 'FOUND' : 'NOT FOUND'));
  if (profile) Logger.log('Profile data: ' + JSON.stringify(profile, null, 2));
  return profile;
}

function clearUserProfile() {
  const userEmail = Session.getActiveUser().getEmail();
  PropertiesService.getUserProperties().deleteAllProperties();
  CacheService.getUserCache().remove('user_profile_' + userEmail);
  Logger.log('✅ Profile cleared for: ' + userEmail);
}

function adminGetTrackingSheetStats() {
  try {
    const trackingSheet = SpreadsheetApp.openById(USER_TRACKING_SHEET_ID);
    const sheet = trackingSheet.getSheetByName('Users');
    if (!sheet) return { success: false, error: 'Tracking sheet not found' };
    const data      = sheet.getDataRange().getValues();
    const userCount = data.length - 1;
    Logger.log('Total users: ' + userCount);
    data.slice(1).forEach(row => Logger.log(`- ${row[0]} (${row[1]}) - Sheet: ${row[2]}`));
    return {
      success:    true,
      totalUsers: userCount,
      users:      data.slice(1).map(row => ({ firstName: row[0], email: row[1], spreadsheetId: row[2], theme: row[4], creationDate: row[11] }))
    };
  } catch (e) {
    Logger.log('ERROR: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}
