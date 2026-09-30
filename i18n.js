(function exposePartsListI18n(root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  if (root) root.PartsListI18n = api;
}(typeof globalThis !== 'undefined' ? globalThis : this, () => {
  'use strict';

  // The interface is written in English. A language pack is a list of [english, translation] pairs.
  // Pairs without placeholders are exact matches. Pairs with {0}, {1}… are templates written against the
  // text as it is rendered, e.g. ['{0} parts selected', '{0} artiklar valda'].
  const PLACEHOLDER = /\{(\d+)\}/g;
  const HAS_PLACEHOLDER = /\{\d+\}/;
  const SENTENCE_BREAK = /(?<=[.!?])\s+(?=[A-ZÅÄÖ0-9“"(])/;
  const MEMO_LIMIT = 6000;

  const packs = new Map();

  // ISO 3166-1 numeric → alpha-2, used to name countries in the reader's language.
  const ISO_NUMERIC = '004AF 008AL 010AQ 012DZ 016AS 020AD 024AO 028AG 031AZ 032AR 036AU 040AT 044BS 048BH 050BD 051AM 052BB 056BE 060BM 064BT 068BO 070BA 072BW 076BR 084BZ 086IO 090SB 092VG 096BN 100BG 104MM 108BI 112BY 116KH 120CM 124CA 132CV 136KY 140CF 144LK 148TD 152CL 156CN 158TW 170CO 174KM 178CG 180CD 184CK 188CR 191HR 192CU 196CY 203CZ 204BJ 208DK 212DM 214DO 218EC 222SV 226GQ 231ET 232ER 233EE 234FO 238FK 239GS 242FJ 246FI 248AX 250FR 258PF 260TF 262DJ 266GA 268GE 270GM 275PS 276DE 288GH 296KI 300GR 304GL 308GD 316GU 320GT 324GN 328GY 332HT 334HM 336VA 340HN 344HK 348HU 352IS 356IN 360ID 364IR 368IQ 372IE 376IL 380IT 384CI 388JM 392JP 398KZ 400JO 404KE 408KP 410KR 414KW 417KG 418LA 422LB 426LS 428LV 430LR 434LY 438LI 440LT 442LU 446MO 450MG 454MW 458MY 462MV 466ML 470MT 478MR 480MU 484MX 492MC 496MN 498MD 499ME 500MS 504MA 508MZ 512OM 516NA 520NR 524NP 528NL 531CW 533AW 534SX 540NC 548NH 554NZ 558NI 562NE 566NG 570NU 574NF 578NO 580MP 583FM 584MH 585PW 586PK 591PA 598PG 600PY 604PE 608PH 612PN 616PL 620PT 624GW 626TL 630PR 634QA 642RO 643RU 646RW 652BL 654SH 659KN 660AI 662LC 663MF 666PM 670VC 674SM 678ST 682SA 686SN 688RS 690SC 694SL 702SG 703SK 704VD 705SI 706SO 710ZA 716RH 724ES 728SS 729SD 732EH 740SR 748SZ 752SE 756CH 760SY 762TJ 764TH 768TG 776TO 780TT 784AE 788TN 792TR 795TM 796TC 800UG 804UA 807MK 818EG 826GB 831GG 832JE 833IM 834TZ 840US 850VI 854BF 858UY 860UZ 862VE 876WF 882WS 887YE 894ZM'
    .split(' ')
    .reduce((map, entry) => { map[entry.slice(0, 3)] = entry.slice(3); return map; }, {});

  const displayNames = new Map();

  function escapeRegExp(text) {
    return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function compileTemplate(source, target) {
    const literals = source.split(PLACEHOLDER).filter((_, index) => index % 2 === 0);
    const order = [...source.matchAll(PLACEHOLDER)].map((match) => Number(match[1]));
    let pattern = '^';
    literals.forEach((literal, index) => {
      pattern += escapeRegExp(literal);
      if (index < order.length) pattern += '([\\s\\S]*?)';
    });
    pattern += '$';
    return {
      regex: new RegExp(pattern),
      order,
      target,
      literals: literals.filter(Boolean),
      // Placeholders of a template without its own " · " must not reach across the separators of a list.
      crossesLists: source.includes(' · '),
      weight: literals.join('').length,
    };
  }

  // English → translated country names for every country in the table, from the browser's own data.
  function countryPairs(language) {
    const pairs = [];
    try {
      const english = new Intl.DisplayNames(['en'], { type: 'region' });
      const target = new Intl.DisplayNames([language], { type: 'region' });
      Object.values(ISO_NUMERIC).forEach((alpha2) => {
        const from = english.of(alpha2);
        const to = target.of(alpha2);
        if (from && to && from !== alpha2 && to !== alpha2 && from !== to) pairs.push([from, to]);
      });
    } catch { /* Country names are optional. */ }
    return pairs;
  }

  function register(language, pairs) {
    const exact = new Map();
    const templates = [];
    pairs.forEach(([source, target]) => {
      if (HAS_PLACEHOLDER.test(source)) templates.push(compileTemplate(source, target));
      else exact.set(source, target);
    });
    // The most specific template (most fixed text) is tried first.
    // Explicit pairs win over generated country names.
    if (language !== 'en') countryPairs(language).forEach(([source, target]) => { if (!exact.has(source)) exact.set(source, target); });
    templates.sort((a, b) => b.weight - a.weight);
    packs.set(language, { exact, templates, memo: new Map() });
  }

  // A captured value may itself be translatable, e.g. a country or a unit word.
  function translateCapture(pack, value, depth) {
    const core = value.trim();
    if (!core || depth >= 3) return value;
    const hit = lookup(pack, core, depth + 1) ?? translateParts(pack, core);
    return hit === undefined ? value : value.replace(core, hit);
  }

  function matchTemplate(pack, text, depth) {
    for (const template of pack.templates) {
      // Cheap pre-check before running the regular expression.
      if (!template.literals.every((literal) => text.includes(literal))) continue;
      const match = template.regex.exec(text);
      if (!match) continue;
      if (!template.crossesLists && match.slice(1).some((value) => value.includes(' · '))) continue;
      const result = template.target.replace(PLACEHOLDER, (_, number) => {
        const index = template.order.indexOf(Number(number));
        return index >= 0 ? translateCapture(pack, match[index + 1], depth) : '';
      });
      // A template that leaves the text unchanged (a generic "{0} → {1}") is not a translation.
      if (result !== text) return result;
    }
    return undefined;
  }

  // Whole-string lookup: exact match first, then templates.
  function lookup(pack, text, depth = 0) {
    if (pack.exact.has(text)) return pack.exact.get(text);
    if (!/[A-Za-zÅÄÖåäö]{2}/.test(text)) return undefined;
    return matchTemplate(pack, text, depth);
  }

  // Messages built from several parts are translated part by part.
  function translateParts(pack, text) {
    const sentences = text.split(SENTENCE_BREAK);
    if (sentences.length > 1) {
      let changed = false;
      const parts = sentences.map((sentence) => {
        const hit = lookup(pack, sentence) ?? translateBullets(pack, sentence);
        if (hit !== undefined) changed = true;
        return hit ?? sentence;
      });
      if (changed) return parts.join(' ');
    }
    return translateBullets(pack, text);
  }

  // "a · b · c" lists: at each position take the longest run of segments that is a known phrase, so
  // "Portugal · EU · EEA agreement" finds "EU · EEA agreement" instead of three separate words.
  function translateBullets(pack, text) {
    if (!text.includes(' · ')) return undefined;
    const segments = text.split(' · ');
    const output = [];
    let changed = false;
    for (let start = 0; start < segments.length;) {
      let matched = false;
      for (let end = segments.length; end > start; end -= 1) {
        const phrase = segments.slice(start, end).join(' · ');
        // A run of several segments must be a known phrase; templates only apply to one segment, otherwise a
        // placeholder could swallow its neighbours ("{0} source prices" matching "Import · USD source prices").
        const hit = end - start === 1 ? lookup(pack, phrase) : pack.exact.get(phrase);
        if (hit === undefined) continue;
        output.push(hit);
        start = end;
        matched = true;
        changed = true;
        break;
      }
      if (!matched) {
        output.push(segments[start]);
        start += 1;
      }
    }
    return changed ? output.join(' · ') : undefined;
  }

  // "Apply for Tollkreditt." is the known phrase "Apply for Tollkreditt" plus its full stop.
  function translateTrailingPunctuation(pack, text) {
    const match = /^([\s\S]*?[A-Za-z0-9)”"])([.:;!…]+)$/.exec(text);
    if (!match) return undefined;
    const hit = lookup(pack, match[1]) ?? translateBullets(pack, match[1]);
    return hit === undefined ? undefined : hit + match[2];
  }

  function translate(language, text) {
    const pack = packs.get(language);
    if (!pack || typeof text !== 'string' || !text) return text;
    if (!/[A-Za-z]{2}/.test(text)) return text;
    if (pack.memo.has(text)) return pack.memo.get(text);
    const leading = text.match(/^\s*/)[0];
    const trailing = text.match(/\s*$/)[0];
    const core = text.trim();
    const hit = lookup(pack, core) ?? translateParts(pack, core) ?? translateTrailingPunctuation(pack, core);
    const result = hit === undefined ? text : `${leading}${hit}${trailing}`;
    if (pack.memo.size >= MEMO_LIMIT) pack.memo.clear();
    pack.memo.set(text, result);
    return result;
  }

  // Adds exact pairs after registration, e.g. country names that only become known once the map loads.
  function addPairs(language, pairs) {
    const pack = packs.get(language);
    if (!pack) return;
    pairs.forEach(([source, target]) => {
      if (source && target && source !== target && !pack.exact.has(source)) pack.exact.set(source, target);
    });
    pack.memo.clear();
  }

  function hasTranslation(language, text) {
    return translate(language, text) !== text;
  }

  function countryName(language, numericId, fallback = '') {
    const alpha2 = ISO_NUMERIC[String(numericId)];
    if (!alpha2 || language === 'en') return fallback;
    try {
      if (!displayNames.has(language)) displayNames.set(language, new Intl.DisplayNames([language], { type: 'region' }));
      const name = displayNames.get(language).of(alpha2);
      return name && name !== alpha2 ? name : fallback;
    } catch {
      return fallback;
    }
  }

  function locale(language) {
    return language === 'sv' ? 'sv-SE' : 'en-GB';
  }

  return { register, addPairs, translate, hasTranslation, countryName, locale, ISO_NUMERIC };
}));
