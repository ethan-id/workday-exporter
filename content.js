// ===== Config =====
const TZ = 'America/Chicago';

// ===== Regex & helpers =====
const TIME_RX = /\b\d{1,2}:\d{2}\s*[AP]M\s*[-–]\s*\d{1,2}:\d{2}\s*[AP]M\b/i;
const MDY_RX = /\b(\d{1,2})\/(\d{1,2})\/(\d{4})\b/;

const IDS = {
    COURSE: '56$380280', // Course Listing
    MP: '56$532392', // Meeting Patterns
    START: '56$435880', // Start Date
    END: '56$435879' // End Date
};

function s(x) {
    return (x || '')
        .replace(/\u00A0/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}
function icsEscape(t) {
    return s(t).replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\n/g, '\\n');
}
function yyyymmdd(d) {
    const y = d.getFullYear(),
        m = String(d.getMonth() + 1).padStart(2, '0'),
        dd = String(d.getDate()).padStart(2, '0');
    return `${y}${m}${dd}`;
}
function hhmmss(d) {
    return `${String(d.getHours()).padStart(2, '0')}${String(d.getMinutes()).padStart(2, '0')}00`;
}
function dtstampNowUTC() {
    const d = new Date();
    return `${d.getUTCFullYear()}${String(d.getUTCMonth() + 1).padStart(2, '0')}${String(d.getUTCDate()).padStart(2, '0')}T${String(d.getUTCHours()).padStart(2, '0')}${String(d.getUTCMinutes()).padStart(2, '0')}${String(d.getUTCSeconds()).padStart(2, '0')}Z`;
}

function parseTime12h(s12) {
    const m = s12.match(/^(\d{1,2}):(\d{2})\s*([AP]M)$/i);
    if (!m) return null;
    let h = +m[1];
    const min = +m[2];
    const ampm = m[3].toUpperCase();
    if (ampm === 'PM' && h !== 12) h += 12;
    if (ampm === 'AM' && h === 12) h = 0;
    return {h, m: min};
}
function parseMDY(s1) {
    const m = s1.match(MDY_RX);
    if (!m) return null;
    return {month: +m[1], day: +m[2], year: +m[3]};
}
function asLocalDate(y, mo, d, h = 0, min = 0) {
    return new Date(y, mo - 1, d, h, min, 0, 0);
}
function nextWeekdayOnOrAfter(start, weekday0Sun) {
    const d = new Date(start);
    const delta = (weekday0Sun - d.getDay() + 7) % 7;
    d.setDate(d.getDate() + delta);
    return d;
}

// ===== Days parsing =====
const DAY_MAP = {
    M: 'MO',
    T: 'TU',
    W: 'WE',
    R: 'TH',
    Th: 'TH',
    Tu: 'TU',
    F: 'FR',
    Sa: 'SA',
    Sat: 'SA',
    Su: 'SU',
    Sun: 'SU'
};
function parseDaysToken(daysRaw) {
    let t = (daysRaw || '').replace(/\s+/g, '');
    const out = [];
    for (let i = 0; i < t.length; ) {
        const two = t.slice(i, i + 2),
            three = t.slice(i, i + 3).toLowerCase();
        if (two === 'Th') {
            out.push('Th');
            i += 2;
            continue;
        }
        if (two === 'Tu') {
            out.push('Tu');
            i += 2;
            continue;
        }
        if (three === 'sat') {
            out.push('Sat');
            i += 3;
            continue;
        }
        if (three === 'sun') {
            out.push('Sun');
            i += 3;
            continue;
        }
        out.push(t[i]);
        i += 1;
    }
    return out.map((tok) => DAY_MAP[tok] || DAY_MAP[tok[0]]).filter(Boolean);
}

// ===== MP parsing from "MWF | 12:05 PM - 12:55 PM | Room ..." =====
function splitChunks(text) {
    return text
        .split(/\n|;|•|·|\s*\|\s*/g)
        .map(s)
        .filter(Boolean);
}

// Replace your existing parseMeetingPattern() with this version
function parseMeetingPattern(raw) {
    const text = s(raw);
    const parts = splitChunks(text); // splits on pipes, bullets, newlines, etc.

    let days = [],
        startTime = null,
        endTime = null,
        location = '';

    // Find the first chunk that looks like a time range
    const ti = parts.findIndex((p) => TIME_RX.test(p));

    if (ti !== -1) {
        // 1) Parse times from the time-chunk
        const timeChunk = parts[ti];
        const m = timeChunk.match(/(\d{1,2}:\d{2}\s*[AP]M)\s*[-–]\s*(\d{1,2}:\d{2}\s*[AP]M)/i);
        if (m) {
            startTime = parseTime12h(m[1].toUpperCase().replace(/\s+/g, ' '));
            endTime = parseTime12h(m[2].toUpperCase().replace(/\s+/g, ' '));
        }

        // 2) Days normally live in the chunk BEFORE the time-chunk
        if (ti - 1 >= 0) {
            days = parseDaysToken(parts[ti - 1] || '');
        }

        // 3) Fallback: if no days yet, they might be glued to the SAME time chunk
        //    e.g. "MWF 12:05 PM - 12:55 PM" or "TR 9:30 AM - 10:45 AM"
        if (!days.length) {
            const beforeTime = timeChunk.split(/\d{1,2}:\d{2}\s*[AP]M/i)[0]; // text before the first time
            const maybeDays = s(beforeTime);
            if (maybeDays) {
                const parsed = parseDaysToken(maybeDays);
                if (parsed.length) days = parsed;
            }
        }

        // 4) Location is everything AFTER the time-chunk
        location = parts.slice(ti + 1).join(' ');
    } else {
        // Fallback: free-form lines like "MoWeFr 11:00 AM - 11:50 AM • Howe Hall 1244"
        const t = text.match(/(\d{1,2}:\d{2}\s*[AP]M)\s*[-–]\s*(\d{1,2}:\d{2}\s*[AP]M)/i);
        if (t) {
            startTime = parseTime12h(t[1].toUpperCase().replace(/\s+/g, ' '));
            endTime = parseTime12h(t[2].toUpperCase().replace(/\s+/g, ' '));
        }
        const m = text.match(/^([A-Za-z]{1,12})\s+\d{1,2}:/);
        if (m) days = parseDaysToken(m[1]);
        // location: chunk after time
        const chunks = text.split(/•|\u2022|·|\s\|\s/).map(s);
        const idx = chunks.findIndex((c) => TIME_RX.test(c));
        if (idx >= 0 && chunks[idx + 1]) location = chunks[idx + 1];
    }

    // Optional date range embedded in the raw string
    let startDate = null,
        endDate = null;
    const dr = text.match(/(\d{1,2}\/\d{1,2}\/\d{4})\s*[-–]\s*(\d{1,2}\/\d{1,2}\/\d{4})/);
    if (dr) {
        const s1 = parseMDY(dr[1]);
        const s2 = parseMDY(dr[2]);
        if (s1) startDate = asLocalDate(s1.year, s1.month, s1.day);
        if (s2) endDate = asLocalDate(s2.year, s2.month, s2.day);
    }

    return {days, startTime, endTime, location: s(location), startDate, endDate};
}

function expandOccurrences(mp, courseTitle) {
    const out = [];
    if (!mp.days.length || !mp.startTime || !mp.endTime) return out;

    const startDate =
        mp.startDate ||
        (() => {
            const d = new Date();
            return nextWeekdayOnOrAfter(d, 1);
        })();
    const endDate =
        mp.endDate ||
        (() => {
            const d = new Date(startDate);
            d.setDate(d.getDate() + 7 * 16);
            return d;
        })();

    for (const byday of mp.days) {
        const weekday0Sun = {SU: 0, MO: 1, TU: 2, WE: 3, TH: 4, FR: 5, SA: 6}[byday];
        if (weekday0Sun == null) continue;
        let cur = nextWeekdayOnOrAfter(startDate, weekday0Sun);
        while (cur <= endDate) {
            const st = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate(), mp.startTime.h, mp.startTime.m, 0, 0);
            const en = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate(), mp.endTime.h, mp.endTime.m, 0, 0);
            if (en <= st) en.setMinutes(en.getMinutes() + 50);
            out.push({title: courseTitle, location: mp.location, start: st, end: en});
            cur.setDate(cur.getDate() + 7);
        }
    }
    return out;
}

// ===== DOM helpers =====
function getCoursesTable() {
    const tables = [...document.querySelectorAll('table')];
    for (const t of tables) {
        const cap = s(t.querySelector('caption')?.textContent);
        if (/^My\s+Enrolled\s+Courses$/i.test(cap)) return t;
    }
    for (const t of tables) {
        const cap = s(t.querySelector('caption')?.textContent);
        if (/My\s+Enrolled\s+Courses/i.test(cap)) return t;
    }
    return null;
}

function visibleText(el) {
    if (!el) return '';
    let txt = s(el.innerText);
    if (txt) return txt;
    const bits = [];
    el.querySelectorAll('[aria-label],[title]').forEach((n) => {
        const v = n.getAttribute('aria-label') || n.getAttribute('title');
        if (v) bits.push(v);
    });
    return s(bits.join(' '));
}

// Course title from the COURSE metadata cell; fallback to any non-empty cell
function getCourseTitleFromRow(row) {
    const courseCell = row.querySelector(`[data-metadata-id="${IDS.COURSE}"]`);
    const course = visibleText(courseCell);
    if (course) return course;
    const first = [...row.querySelectorAll('td')].map((td) => s(td.innerText)).find(Boolean);
    return first || 'Course';
}

// Meeting patterns from MP metadata cell (prefer aria/title carrying full "DAYS | TIME | LOCATION")
function getMeetingStringFromRow(row) {
    const mpCell = row.querySelector(`[data-metadata-id="${IDS.MP}"]`);
    if (mpCell) {
        const attr = [...mpCell.querySelectorAll('[aria-label],[title]')]
            .map((n) => n.getAttribute('aria-label') || n.getAttribute('title'))
            .map(s)
            .find((v) => TIME_RX.test(v));
        if (attr) return attr;
        const txt = visibleText(mpCell);
        if (TIME_RX.test(txt)) return txt;
    }
    // Fallback: any visible text with a time range
    const any = [...row.querySelectorAll('a,div,span')].map(visibleText).find((v) => TIME_RX.test(v));
    return any || '';
}

// Row-level start/end dates from their metadata cells; fallback: scan for MDY
function getStartEndDatesFromRow(row) {
    const sCell = row.querySelector(`[data-metadata-id="${IDS.START}"]`);
    const eCell = row.querySelector(`[data-metadata-id="${IDS.END}"]`);
    const sTxt = s(sCell?.innerText || sCell?.getAttribute('title') || '');
    const eTxt = s(eCell?.innerText || eCell?.getAttribute('title') || '');

    let start = null,
        end = null;
    if (MDY_RX.test(sTxt)) {
        const m = parseMDY(sTxt);
        if (m) start = asLocalDate(m.year, m.month, m.day);
    }
    if (MDY_RX.test(eTxt)) {
        const m = parseMDY(eTxt);
        if (m) end = asLocalDate(m.year, m.month, m.day);
    }

    if (!start || !end) {
        const texts = [...row.querySelectorAll('td div, td a, td span')].map((el) =>
            s(el.textContent || el.getAttribute('title') || '')
        );
        const mdy = texts.filter((t) => MDY_RX.test(t));
        if (!start && mdy[0]) {
            const m = parseMDY(mdy[0]);
            start = asLocalDate(m.year, m.month, m.day);
        }
        if (!end && mdy[mdy.length - 1]) {
            const m = parseMDY(mdy[mdy.length - 1]);
            end = asLocalDate(m.year, m.month, m.day);
        }
    }
    return {start, end};
}

// ===== Build ICS =====
function toICS(events) {
    const lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'CALSCALE:GREGORIAN', 'PRODID:-//ISU Workday Export//EN'];
    const stamp = dtstampNowUTC();
    events.forEach((ev, i) => {
        const uid = `isu-workday-${i}-${ev.start.getTime()}@isu`;
        lines.push(
            'BEGIN:VEVENT',
            `UID:${uid}`,
            `DTSTAMP:${stamp}`,
            `SUMMARY:${icsEscape(ev.title)}`,
            `LOCATION:${icsEscape(ev.location || '')}`,
            `DTSTART;TZID=${TZ}:${yyyymmdd(ev.start)}T${hhmmss(ev.start)}`,
            `DTEND;TZID=${TZ}:${yyyymmdd(ev.end)}T${hhmmss(ev.end)}`,
            'END:VEVENT'
        );
    });
    lines.push('END:VCALENDAR');
    return lines.join('\r\n');
}
function downloadICS(filename, icsText) {
    const blob = new Blob([icsText], {type: 'text/calendar;charset=utf-8'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
        URL.revokeObjectURL(url);
        a.remove();
    }, 0);
}

// ===== Main scrape =====
function scrapeCourses() {
    const table = getCoursesTable();
    if (!table) throw new Error("Could not find the 'My Enrolled Courses' table.");

    const rows = [...table.querySelectorAll('tbody tr')].filter((r) => r.querySelector('td'));
    const events = [];

    for (const row of rows) {
        const courseTitle = getCourseTitleFromRow(row); // via 56$380280
        const mpString = getMeetingStringFromRow(row); // via 56$532392
        if (!mpString) continue;

        const {start: startDate, end: endDate} = getStartEndDatesFromRow(row); // 56$435880 / 56$435879

        // Reconstruct "DAYS TIME • LOCATION" from pipe-separated MP
        const parts = splitChunks(mpString);
        const ti = parts.findIndex((p) => TIME_RX.test(p));
        const candidates = [];
        if (ti !== -1) {
            const days = parts[ti - 1] || '';
            const time = parts[ti];
            const loc = parts.slice(ti + 1).join(' ');
            candidates.push(`${days} ${time}${loc ? ` • ${loc}` : ''}`);
        } else {
            candidates.push(mpString);
        }

        for (const cand of candidates) {
            const mp = parseMeetingPattern(cand);
            if (!mp.days.length || !mp.startTime || !mp.endTime) continue;

            // If MP lacked dates, apply row-level dates
            if (!mp.startDate && startDate) mp.startDate = startDate;
            if (!mp.endDate && endDate) mp.endDate = endDate;

            const evs = expandOccurrences(mp, courseTitle);
            if (evs.length) events.push(...evs);
        }
    }

    return events;
}

function runExport() {
    const events = scrapeCourses();
    console.info('[ISU Export] Parsed events:', events.length);
    if (!events.length) {
        alert('No class meetings found to export. Make sure the table is visible and expanded.');
        return;
    }
    const termGuess = (() => {
        const t = document.title || '';
        const m = t.match(/(\bFall|\bSpring|\bSummer)\s+(\d{4})/i);
        return m ? `${m[1]} ${m[2]}` : 'courses';
    })();
    const ics = toICS(events);
    downloadICS(`ISU-${termGuess}.ics`, ics);
}

// Bridge from popup
chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    if (msg?.type === 'EXPORT_ICS') {
        try {
            runExport();
            sendResponse({ok: true});
        } catch (e) {
            console.error(e);
            alert(e?.message || 'Export failed.');
            sendResponse({ok: false, error: e?.message});
        }
    }
});
