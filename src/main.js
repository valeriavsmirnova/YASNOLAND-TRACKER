import './styles.css';
import * as XLSX from 'xlsx';

const PUBLIC_KEY = 'https://disk.yandex.ru/d/AH9oiy7YcSWXtA';
const API = `https://cloud-api.yandex.net/v1/disk/public/resources/download?public_key=${encodeURIComponent(PUBLIC_KEY)}&path=${encodeURIComponent('/yasnoland-tracker.xlsx')}`;
const LOCAL_FILE = `${import.meta.env.BASE_URL}yasnoland-tracker.xlsx`;
const app = document.querySelector('#app');
let db = null;
let currentStudent = null;
let currentView = 'login';
const LOGIN_KEY = 'yasnoland_remembered_student';
const LANGUAGE_KEY = 'yasnoland_language';
let uiLanguage = (() => { try { return localStorage.getItem(LANGUAGE_KEY) || 'en'; } catch { return 'en'; } })();
const copy = {
  en: { rewards:'Rewards', progress:'Progress', exit:'Exit' },
  ru: { rewards:'Награды', progress:'Прогресс', exit:'Выйти' }
};
const t = key => copy[uiLanguage]?.[key] || copy.en[key] || key;

function rememberedStudentId() {
  try { return localStorage.getItem(LOGIN_KEY) || ''; } catch { return ''; }
}

function rememberStudent(id) {
  try { localStorage.setItem(LOGIN_KEY, id); } catch { /* Storage may be disabled. */ }
}

function forgetStudent() {
  try { localStorage.removeItem(LOGIN_KEY); } catch { /* Storage may be disabled. */ }
}

const esc = (value = '') => String(value).replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
const norm = value => String(value ?? '').trim();
const num = value => Number(value) || 0;
const percent = value => Math.max(0, Math.min(100, Number.parseFloat(String(value ?? '').replace(',', '.')) || 0));
const get = (row, ...keys) => { const found = Object.keys(row).find(k => keys.some(key => k.trim().toLowerCase() === key.toLowerCase())); return found ? row[found] : ''; };
const statusClass = status => norm(status).toLowerCase().replaceAll(' ', '-');

function workbookToDb(workbook) {
  const out = {};
  workbook.SheetNames.forEach(name => {
    out[name] = XLSX.utils.sheet_to_json(workbook.Sheets[name], { defval: '', raw: false })
      .filter(row => Object.values(row).some(value => norm(value)));
  });
  return out;
}

async function fetchWorkbook() {
  let remoteError;
  try {
    const apiResponse = await fetch(API, { cache: 'no-store' });
    if (!apiResponse.ok) throw new Error(`Yandex API: ${apiResponse.status}`);
    const { href } = await apiResponse.json();
    if (!href) throw new Error('Yandex API did not return a download link');
    const fileResponse = await fetch(href, { cache: 'no-store' });
    if (!fileResponse.ok) throw new Error(`Excel download: ${fileResponse.status}`);
    const buffer = await fileResponse.arrayBuffer();
    if (buffer.byteLength < 1000) throw new Error('Downloaded Excel file is unexpectedly small');
    return { buffer, source: 'Yandex Disk' };
  } catch (error) { remoteError = error; }
  const localResponse = await fetch(LOCAL_FILE, { cache: 'no-store' });
  if (!localResponse.ok) throw new Error(`${remoteError?.message}; local fallback: ${localResponse.status}`);
  return { buffer: await localResponse.arrayBuffer(), source: 'local fallback', warning: remoteError?.message };
}

async function loadData(showNotice = false) {
  setNotice('Loading data...', 'loading');
  try {
    const result = await fetchWorkbook();
    db = workbookToDb(XLSX.read(result.buffer, { type: 'array', cellDates: true }));
    if (!db.Students?.length) throw new Error('Students sheet is empty');
    setNotice(result.warning ? 'Data updated · local fallback' : 'Data updated', result.warning ? 'warning' : 'success');
    if (currentStudent) {
      currentStudent = findStudent(currentStudent.id);
      currentView === 'progress' ? renderProgress() : renderDashboard();
    } else if (currentView === 'leaderboard') renderLeaderboard();
    else {
      const remembered = findStudent(rememberedStudentId());
      if (remembered?.status.toLowerCase() === 'active') { currentStudent = remembered; renderDashboard(); }
      else { if (rememberedStudentId()) forgetStudent(); renderLogin(); }
    }
  } catch (error) {
    console.error(error);
    setNotice('Unable to load data', 'error');
    if (!db) renderError(error.message);
  }
}

function studentView(row) {
  return {
    id: norm(get(row, 'Student ID')),
    first: norm(get(row, 'Имя', 'First Name')),
    last: norm(get(row, 'Фамилия', 'Last Name')),
    group: norm(get(row, 'Группа', 'Group')),
    course: norm(get(row, 'Курс', 'Course')),
    pin: norm(get(row, 'Код доступа', 'Code / PIN', 'PIN / Access Code')),
    status: norm(get(row, 'Статус', 'Status'))
  };
}

function findStudent(id) {
  const row = db?.Students?.find(item => norm(get(item, 'Student ID')).toLowerCase() === norm(id).toLowerCase());
  return row ? studentView(row) : null;
}

function shell(content, dashboard = false, publicView = false) {
  app.innerHTML = `<header class="topbar ${dashboard ? 'student-topbar' : ''}"><a class="brand" href="#" aria-label="YASNOLAND"><span class="brand-mark">Y</span><span><b>YASNO<span>LAND</span></b><small>ENGLISH BECOMES CLEAR</small></span></a>${dashboard ? dashboardTabs(currentView === 'progress' ? 'progress' : 'rewards') : ''}<div class="top-actions">${dashboard ? `<label class="language-picker"><span>${uiLanguage === 'en' ? '🇬🇧' : '🇷🇺'}</span><select aria-label="Language"><option value="en" ${uiLanguage === 'en' ? 'selected' : ''}>English</option><option value="ru" ${uiLanguage === 'ru' ? 'selected' : ''}>Русский</option></select></label><button class="student-chip"><span>${esc(currentStudent.first[0])}</span>${esc(currentStudent.first)}<i>⌄</i></button><button class="logout-link" data-action="logout">${t('exit')}</button>` : publicView ? `<button class="ghost" data-action="login">Student login</button>` : ''}${!dashboard ? `<button class="refresh" data-action="refresh"><svg viewBox="0 0 24 24"><path d="M20 6v5h-5M4 18v-5h5M6.1 9a7 7 0 0 1 11.4-2.5L20 9M4 15l2.5 2.5A7 7 0 0 0 17.9 15"/></svg>Refresh data</button>` : ''}</div></header>${content}<div id="notice" class="notice" aria-live="polite"></div>`;
  const earningTitle = document.querySelector('.hero-earning-rules > p');
  if (earningTitle) earningTitle.textContent = uiLanguage === 'ru' ? 'Как заработать доллары на каждом уроке' : 'How to earn dollars every lesson';
  bindActions();
  localizePage();
  updateMonthlySummary();
}

function updateMonthlySummary() {
  const card = document.querySelector('.progress-month-card.active');
  const values = [...(card?.querySelectorAll('.academic-bar b') || [])].map(item => Number.parseFloat(item.textContent)).filter(Number.isFinite);
  const score = values.length ? Math.round(values.reduce((sum,value) => sum + value, 0) / values.length) : 0;
  const summary = document.querySelector('.mobile-year-average');
  if (!summary) return;
  const label = summary.querySelector('small');
  const result = summary.querySelector('strong');
  if (label) label.textContent = uiLanguage === 'ru' ? 'Общий результат за месяц' : 'Overall result for this month';
  if (result) result.textContent = `${score}%`;
}

function localizePage() {
  document.documentElement.lang = uiLanguage === 'ru' ? 'ru' : 'en';
  if (uiLanguage !== 'ru') return;
  const translations = new Map(Object.entries({
    'YOUR YASNOLAND SPACE':'ТВОЙ ЯСНОЛЕНД', 'MY BALANCE':'МОЙ БАЛАНС', 'YASNOLAND DOLLARS':'ДОЛЛАРЫ ЯСНОЛЕНДА',
    'How to earn dollars every lesson':'Как заработать доллары на каждом уроке', 'English Only':'Только английский', 'Homework':'Домашнее задание', 'Class Rules':'Правила урока',
    '3 goals':'3 цели', '2 goals':'2 цели', '0–1 goal':'0–1 цель', 'Earn up to 2 YASNOLAND Dollars every lesson.':'За каждый урок можно заработать до 2 долларов Ясноленда.',
    'SPEND YOUR DOLLARS':'ПОТРАТЬ ДОЛЛАРЫ', 'Gift Fair':'Ярмарка подарков', 'Choose something worth learning for.':'Выбери подарок, ради которого хочется стараться.',
    'Available':'Доступно', 'Not available':'Недоступно', 'WELCOME BONUS':'ПРИВЕТСТВЕННЫЙ БОНУС', 'Secret Gift':'Секретный подарок',
    'Save your first 3\nYASNOLAND DOLLARS':'Сохрани первые 3\nДОЛЛАРА ЯСНОЛЕНДА', '“Exchange them at the Gift Fair at the end of the school year!”':'«Обменяй их на ярмарке подарков в конце учебного года!»',
    'MY DOLLARS':'МОИ ДОЛЛАРЫ', 'See how your balance grows':'Как растёт твой баланс', 'See my lessons':'Мои уроки', 'YASNOLAND DOLLARS EARNED':'ЗАРАБОТАНО ДОЛЛАРОВ ЯСНОЛЕНДА',
    'Lessons':'Уроки', 'Completed':'Завершено', 'In progress':'В процессе', 'ACADEMIC PROGRESS':'УЧЕБНЫЙ ПРОГРЕСС', 'Overall Academic Progress':'Общий учебный прогресс',
    'LATEST MONTH':'ПОСЛЕДНИЙ МЕСЯЦ', 'MONTH RESULT':'РЕЗУЛЬТАТ МЕСЯЦА', 'MONTHLY ACADEMIC PROGRESS':'ПРОГРЕСС ПО МЕСЯЦАМ',
    'Progress Tracker':'Трекер прогресса', 'Average':'Среднее', 'for all months':'за все месяцы', 'Skills and homework':'Навыки и домашнее задание',
    'Vocabulary':'Словарный запас', 'Grammar':'Грамматика', 'Reading':'Чтение', 'Listening':'Аудирование', 'Speaking':'Говорение',
    'September':'Сентябрь', 'October':'Октябрь', 'November':'Ноябрь', 'December':'Декабрь', 'January':'Январь', 'February':'Февраль', 'March':'Март', 'April':'Апрель', 'May':'Май',
    'Small gift':'Маленький подарок', 'Stationery':'Канцелярия', 'Special':'Особенный подарок', 'Dollars':'Долларов', 'Yasnoland Dollars':'Доллары Ясноленда',
    "Teacher's Comment":'Комментарий преподавателя', 'Progress data is not available yet.':'Данные о прогрессе пока недоступны.',
    'See how your Dollars grow':'Как растут твои доллары', 'Your rewards':'Твои подарки', 'MY PURCHASES':'МОИ ПОКУПКИ'
  }));
  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
  const nodes = []; while (walker.nextNode()) nodes.push(walker.currentNode);
  nodes.forEach(node => {
    const original = node.nodeValue; const trimmed = original.trim(); if (!trimmed) return;
    let translated = translations.get(trimmed);
    if (!translated) translated = trimmed.replace(/September/g,'Сентябрь').replace(/October/g,'Октябрь').replace(/November/g,'Ноябрь').replace(/December/g,'Декабрь').replace(/January/g,'Январь').replace(/February/g,'Февраль').replace(/March/g,'Март').replace(/April/g,'Апрель').replace(/May/g,'Май').replace(/(\d+) lessons/g, '$1 уроков').replace(/(\d+) Dollar to go/g, 'Остался $1 доллар').replace(/(\d+) Dollars to go/g, 'Осталось $1 долларов').replace(/Yasnoland Dollars/g, 'Доллары Ясноленда').replace(/Academic Year/g, 'Учебный год').replace(/ balance$/,' баланс');
    if (translated !== trimmed) node.nodeValue = original.replace(trimmed, translated);
  });
}

function dashboardTabs(active) {
  return `<nav class="dashboard-tabs" aria-label="Student dashboard sections"><button class="${active === 'rewards' ? 'active' : ''}" data-tab="rewards">${t('rewards')}</button><button class="${active === 'progress' ? 'active' : ''}" data-tab="progress">${t('progress')}</button></nav>`;
}

function setNotice(message, type = '') {
  let notice = document.querySelector('#notice');
  if (!notice) { notice = document.createElement('div'); notice.id = 'notice'; notice.className = 'notice'; document.body.append(notice); }
  notice.textContent = message; notice.className = `notice show ${type}`;
  if (type !== 'loading') setTimeout(() => notice?.classList.remove('show'), 3500);
}

function renderLogin() {
  currentView = 'login';
  shell(`<main class="login-page"><section class="login-copy"><p class="eyebrow">REWARDS TRACKER</p><h1>Every lesson is<br><em>a step forward.</em></h1><p>Track your progress, collect YASNOLAND Dollars and choose your next reward.</p><div class="orb orb-one"></div><div class="orb orb-two"></div></section><section class="login-card"><p class="eyebrow">STUDENT ACCESS</p><h2>Welcome back</h2><p class="muted">Enter your personal details to see your progress.</p><form id="login-form"><label>Student ID<input name="id" autocomplete="username" placeholder="For example, Y001" required></label><label>PIN / Access Code<input name="pin" type="password" inputmode="numeric" autocomplete="current-password" placeholder="Your access code" required></label><p id="login-error" class="form-error" role="alert"></p><button class="primary" type="submit">Enter my tracker <span>→</span></button></form><div class="login-divider"><span>or</span></div><button class="leaderboard-button" data-action="leaderboard">View student leaderboard</button><p class="security-note">Client-side access for demonstration purposes. Do not use as secure authentication.</p></section></main><footer>Ясноленд · Где английский становится понятным</footer>`);
  document.querySelector('#login-form').addEventListener('submit', event => {
    event.preventDefault();
    const values = new FormData(event.currentTarget); const student = findStudent(values.get('id'));
    if (!student || student.pin !== norm(values.get('pin')) || student.status.toLowerCase() !== 'active') {
      document.querySelector('#login-error').textContent = 'Student ID or access code is incorrect.'; return;
    }
    currentStudent = student; rememberStudent(student.id); renderDashboard();
  });
}

function renderLeaderboard() {
  currentView = 'leaderboard';
  const rewardRows = db.REWARDS || [];
  const balanceRows = db.BALANCE || [];
  const ranking = (db.Students || []).map(studentView).filter(s => s.status.toLowerCase() === 'active').map(student => {
    const rewards = rewardRows.filter(r => norm(get(r, 'Student ID')) === student.id);
    const balance = balanceRows.find(r => norm(get(r, 'Student ID')) === student.id) || {};
    return { ...student, earned: num(get(balance, 'Earned')), balance: num(get(balance, 'Balance')), completed: rewards.filter(r => norm(get(r, 'Status')).toLowerCase() === 'completed').length };
  }).sort((a, b) => b.earned - a.earned || b.balance - a.balance || a.first.localeCompare(b.first));
  shell(`<main class="leaderboard-page"><section class="leaderboard-hero"><p class="eyebrow">YASNOLAND COMMUNITY</p><h1>Student Leaderboard</h1><p>Общий прогресс активных учеников. Место определяется по заработанным YASNOLAND Dollars.</p></section><section class="leaderboard panel"><div class="leader-head"><span>Place & student</span><span>Course</span><span>Completed</span><span>Earned</span><span>Balance</span></div>${ranking.map((student, index) => `<div class="leader-row ${index < 3 ? `top-${index + 1}` : ''}"><div class="leader-person"><b class="place">${index + 1}</b><span class="avatar">${esc(student.first[0] || 'Y')}${esc(student.last[0] || '')}</span><span><strong>${esc(student.first)} ${esc(student.last)}</strong><small>${esc(student.group)}</small></span></div><span class="leader-course">${esc(student.course)}</span><b>${student.completed} months</b><b class="earned">${student.earned} YD</b><b>${student.balance} YD</b></div>`).join('')}</section><p class="privacy-note">На общем экране не отображаются Student ID и коды доступа.</p></main><footer>Ясноленд · Где английский становится понятным</footer>`, false, true);
}

const star = active => `<svg class="star ${active ? 'active' : ''}" viewBox="0 0 24 24"><path d="m12 2.7 2.8 5.7 6.3.9-4.6 4.4 1.1 6.2-5.6-2.9-5.6 2.9 1.1-6.2-4.6-4.4 6.3-.9Z"/></svg>`;
const score = value => `<span class="stars">${star(num(value) >= 1)}${star(num(value) >= 2)}</span>`;
const metric = (label, value) => `<div class="metric"><span>${esc(label)}</span>${score(value)}</div>`;

const monthNumber = month => ['january','february','march','april','may','june','july','august','september','october','november','december'].indexOf(norm(month).toLowerCase()) + 1;
function lessonDateParts(value) {
  const parts = norm(value).split(/[/.\-]/).map(Number);
  if (parts.length < 3) return null;
  return { month: parts[0], day: parts[1], year: parts[2] < 100 ? 2000 + parts[2] : parts[2] };
}
function lessonDateLabel(value) {
  const date = lessonDateParts(value); if (!date) return norm(value);
  return `${date.day} ${['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][date.month - 1]}`;
}
function goalResult(value) { return num(value) === 1 ? '<span class="goal-done">✓</span>' : '<span class="goal-missed">—</span>'; }
function lessonDollars(value) { const amount = num(value); return `<strong class="lesson-dollars ${amount ? 'earned' : ''}">${amount ? '+' : ''}${amount}</strong>`; }
function rewardMonthCard(reward, studentLessons, index) {
  const month = norm(get(reward,'Month')); const year = num(get(reward,'Year')); const status = norm(get(reward,'Status')); const completed = status.toLowerCase() === 'completed';
  const lessons = num(get(reward,'Lessons')); const dollars = num(get(reward,'Dollars Earned')); const maximum = lessons * 2;
  const monthLessons = studentLessons.filter(lesson => { const date = lessonDateParts(get(lesson,'Date')); return date && date.month === monthNumber(month) && date.year === year; });
  return `<article class="star-month-card"><header><div><small>${esc(month)} ${year}</small><h3>+${dollars}</h3><p>Yasnoland Dollars</p><p>${lessons} lessons · ${dollars} / ${maximum}</p></div><span class="status ${statusClass(status)}">${esc(status)}</span></header><button class="lesson-toggle" data-lessons="lessons-${index}" aria-expanded="false">See my lessons <span>⌄</span></button><div id="lessons-${index}" class="lesson-list" hidden><p class="lesson-list-title">YASNOLAND DOLLARS EARNED</p>${monthLessons.map(lesson => `<div class="lesson-row"><strong>${esc(lessonDateLabel(get(lesson,'Date')))}</strong><span>English Only ${goalResult(get(lesson,'English Only'))}</span><span>Homework ${goalResult(get(lesson,'Homework'))}</span><span>Class Rules ${goalResult(get(lesson,'Class Rules'))}</span>${lessonDollars(get(lesson,'Dollars Earned'))}</div>`).join('') || '<p class="empty-lessons">No lesson details available.</p>'}</div></article>`;
}

function renderDashboard() {
  currentView = 'rewards';
  const id = currentStudent.id;
  const rewards = (db.REWARDS || []).filter(r => norm(get(r, 'Student ID')) === id);
  const visible = rewards.filter(r => norm(get(r, 'Status')).toLowerCase() !== 'not started');
  const current = rewards.find(r => norm(get(r, 'Status')).toLowerCase() === 'in progress') || visible.at(-1) || {};
  const balanceRow = (db.BALANCE || []).find(r => norm(get(r, 'Student ID')) === id) || {};
  const balance = num(get(balanceRow, 'Balance'));
  const academicYear = norm(get(current, 'Year')); const yearLabel = academicYear ? `${academicYear}–${num(academicYear) + 1}` : '';
  const courseLabel = currentStudent.course.replace(/\bbox\b/i, 'Box');
  const purchases = (db.PURCHASES || []).filter(r => norm(get(r, 'Student ID')) === id);
  const studentLessons = (db.LESSONS || []).filter(r => norm(get(r, 'Student ID')) === id);
  const gifts = db.GIFTS || [];
  const balanceStory = visible.length === 2 ? `${get(visible[0],'Month')} +${num(get(visible[0],'Dollars Earned'))} · ${get(visible[1],'Month')} +${num(get(visible[1],'Dollars Earned'))} = ${balance} balance` : 'Each month adds to your balance.';
  shell(`<main class="dashboard rewards-page">${dashboardTabs('rewards')}<section class="welcome"><div><p class="eyebrow">YOUR YASNOLAND SPACE</p><h1><span class="hi-word">Hi,</span> <em>${esc(currentStudent.first)}!</em></h1><p>${esc(courseLabel)}${yearLabel ? ` · ${esc(yearLabel)}` : ''}</p></div></section>
  <section class="rewards-hero"><article class="balance-focus"><div class="balance-copy"><p>MY BALANCE</p><div class="hero-balance">${balance}</div><h2>YASNOLAND<br>DOLLARS</h2><button class="history-button" data-action="history"><svg viewBox="0 0 24 24"><path d="M5 8h14v11H5zM8 8V5h8v3M9 13h6M12 10v6"/></svg>История</button></div><span class="sparkle sparkle-one">✦</span><span class="sparkle sparkle-two">✦</span><img class="wallet-hero" src="${import.meta.env.BASE_URL}images/yasik-wallet-hero.png" alt="YASNOLAND Dollars in a wallet"></article><aside class="hero-earning-rules"><p>HOW TO EARN DOLLARS</p><div class="hero-goals"><span><svg viewBox="0 0 24 24"><path d="M5 6h14v10H9l-4 3zM9 11h.01M12 11h.01M15 11h.01"/></svg>English Only</span><span><svg viewBox="0 0 24 24"><path d="m5 17 1-4L16 3l4 4-10 10zM13 6l4 4M5 20h14"/></svg>Homework</span><span><svg viewBox="0 0 24 24"><path d="m12 3 2.7 5.5 6.1.9-4.4 4.3 1 6.1-5.4-2.9-5.4 2.9 1-6.1-4.4-4.3 6.1-.9z"/></svg>Class Rules</span></div><div class="hero-rule-list"><span class="rule-two"><span class="rule-badge">★</span><b>3 goals</b><span class="mini-note note-two"><img src="${import.meta.env.BASE_URL}images/yasnoland-2-dollar-new.png" alt="2 YASNOLAND Dollars"></span><i>→</i><strong>+2</strong></span><span class="rule-one"><span class="rule-badge">◆</span><b>2 goals</b><span class="mini-note note-one"><img src="${import.meta.env.BASE_URL}images/yasnoland-1-dollar-new.png" alt="1 YASNOLAND Dollar"></span><i>→</i><strong>+1</strong></span><span class="rule-zero"><span class="rule-badge">−</span><b>0–1 goal</b><span class="note-space"></span><i>→</i><strong>0</strong></span></div><small class="earning-hint"><span>i</span>Earn up to 2 YASNOLAND Dollars every lesson.</small></aside></section>
  <section class="gift-section"><div class="section-heading"><div><p class="eyebrow">SPEND YOUR DOLLARS</p><h2>Gift Fair</h2><p class="section-kicker">Choose something worth learning for.</p></div><span class="section-number">01</span></div><div class="gift-grid">${gifts.map(g => giftCard(g,balance)).join('')}<article class="secret-gift-card"><div class="secret-visual"><img src="${import.meta.env.BASE_URL}images/welcome-bonus.png" alt="Ясик с бумажными YASNOLAND Dollars"><span class="yasik-label">Ясик</span><span class="lock-mark"><svg viewBox="0 0 24 24"><path d="M7 10V7a5 5 0 0 1 10 0v3M5 10h14v11H5z"/></svg></span></div><div><p>WELCOME BONUS</p><h3>Secret Gift</h3><strong>Save your first 3<br>YASNOLAND DOLLARS</strong><small class="yasik-message">“Exchange them at the Gift Fair at the end of the school year!”</small></div></article></div></section>
  <section class="dollars-history"><div class="section-heading"><div><p class="eyebrow">MY DOLLARS</p><h2>See how your balance grows</h2></div><div class="balance-story"><span class="section-number">02</span><p>${esc(balanceStory)}</p></div></div><div class="star-months dollars-months">${visible.map((r,i) => rewardMonthCard(r,studentLessons,i)).join('') || '<p class="empty">Earnings will appear here.</p>'}</div></section>
  ${purchases.length ? `<section class="purchases panel"><p class="eyebrow">MY PURCHASES</p><h2>Your rewards</h2><div class="purchase-list">${purchases.map(p => `<div><span>${esc(get(p,'Gift'))}</span><small>${esc(get(p,'Date'))}</small><b>${num(get(p,'Price'))} YD</b></div>`).join('')}</div></section>` : ''}
  </main><footer>Ясноленд · Где английский становится понятным</footer>`, true);
}

const progressSkills = ['Vocabulary', 'Grammar', 'Reading', 'Listening', 'Speaking', 'Homework'];
const academicSkills = progressSkills.slice(0, 5);
const progressTone = value => value >= 100 ? 'high' : value >= 50 ? 'medium' : 'developing';
const progressBar = value => { const score = percent(value); return `<div class="academic-bar ${progressTone(score)}"><i style="width:${score}%"></i><b>${Math.round(score)}%</b></div>`; };

function renderProgress() {
  currentView = 'progress';
  const id = currentStudent.id;
  const rows = (db.PROGRESS || []).filter(row => norm(get(row, 'Student ID')) === id);
  const years = [...new Set(rows.map(row => norm(get(row, 'Year'))).filter(Boolean))].join(' / ');
  if (!rows.length) {
    shell(`<main class="dashboard progress-page">${dashboardTabs('progress')}<section class="progress-identity"><div><p class="eyebrow">ACADEMIC PROGRESS</p><h1>${esc(currentStudent.first)} ${esc(currentStudent.last)}</h1><p>${esc(currentStudent.course)}</p></div></section><section class="no-progress panel"><h2>Progress data is not available yet.</h2><p>Academic results will appear after the teacher adds them to the PROGRESS sheet.</p></section></main><footer>Ясноленд · Где английский становится понятным</footer>`, true);
    return;
  }
  const latest = rows.at(-1);
  const academicValues = rows.flatMap(row => academicSkills.map(skill => get(row, skill))).filter(value => norm(value) !== '').map(percent);
  const academicOverall = academicValues.length ? Math.round(academicValues.reduce((a, b) => a + b, 0) / academicValues.length) : 0;
  const monthRange = rows.length > 1 ? `${get(rows[0], 'Month')}–${get(rows.at(-1), 'Month')}` : norm(get(rows[0], 'Month'));
  const tableHead = rows.map(row => { const comment = norm(get(row, 'Comment')); return `<th><strong>${esc(get(row, 'Month'))}</strong><small>${esc(get(row, 'Unit / Topics'))}</small>${comment ? `<p class="month-comment">${esc(comment)}</p>` : ''}</th>`; }).join('');
  const skillRow = skill => { const values = rows.map(row => norm(get(row, skill)) === '' ? null : percent(get(row, skill))); const available = values.filter(value => value !== null); const overall = available.length ? Math.round(available.reduce((a,b) => a+b,0) / available.length) : 0; return `<tr><th>${skill}</th>${values.map(value => `<td>${value === null ? '<span class="no-value">—</span>' : progressBar(value)}</td>`).join('')}<td class="overall-cell"><b>${overall}%</b></td></tr>`; };
  const tableRows = `<tr class="skill-group"><th colspan="${rows.length + 1}">Skills and homework</th><th class="average-heading"><strong>Average</strong><small>for all months</small></th></tr>${academicSkills.map(skillRow).join('')}${skillRow('Homework')}`;
  const commentedRows = rows.filter(row => norm(get(row, 'Comment')));
  shell(`<main class="dashboard progress-page">${dashboardTabs('progress')}<section class="progress-identity"><div><p class="eyebrow">ACADEMIC PROGRESS</p><h1>${esc(currentStudent.first)} ${esc(currentStudent.last)}</h1><p>${esc(currentStudent.course)} · Academic Year ${esc(years)}</p></div><article class="academic-score"><span>Overall Academic Progress</span><strong>${academicOverall}%</strong><small>${esc(monthRange)}</small><div class="score-ring" style="--score:${academicOverall}"></div></article></section>
  <section class="latest-progress panel"><div><p class="eyebrow">LATEST MONTH</p><h2>${esc(get(latest,'Month'))} <small>${esc(get(latest,'Year'))}</small></h2><p>${esc(get(latest,'Unit / Topics'))}</p></div><div><span>MONTH RESULT</span><strong>${Math.round(percent(get(latest,'Overall Progress')))}%</strong></div></section>
  <section class="progress-table-wrap panel"><div class="progress-section-title"><p class="eyebrow">MONTHLY ACADEMIC PROGRESS</p><h2>Progress Tracker</h2><p class="tracker-intro">Результаты и комментарии преподавателя по каждому месяцу</p></div><table class="progress-table"><thead><tr><th aria-label="Indicator"></th>${tableHead}<th class="average-placeholder" aria-hidden="true"></th></tr></thead><tbody>${tableRows}</tbody></table></section>
  <section class="progress-mobile"><nav class="mobile-month-tabs" aria-label="Months">${rows.map((row,i) => `<button class="${i === rows.length - 1 ? 'active' : ''}" data-mobile-month="${i}">${esc(get(row,'Month'))}</button>`).join('')}</nav>${rows.map((row,i) => { const rowComment = norm(get(row,'Comment')); return `<article class="progress-month-card panel ${i === rows.length - 1 ? 'active' : ''}" data-month-card="${i}"><header><div><h2>${esc(get(row,'Month'))}</h2><p>${esc(get(row,'Unit / Topics'))}</p></div><span class="month-school-mark" aria-hidden="true">⌂</span></header>${rowComment ? `<div class="mobile-comment"><b>Teacher's Comment</b><p>${esc(rowComment)}</p></div>` : ''}<h3 class="mobile-skills-title">Skills and homework</h3>${progressSkills.map(skill => `<div class="mobile-skill"><span>${skill}</span>${progressBar(get(row,skill))}</div>`).join('')}</article>`; }).join('')}<article class="mobile-year-average"><span class="average-wallet">$</span><div><small>Average progress for the year</small><strong>${academicOverall}%</strong><p>Great work! Keep moving forward.</p></div></article></section></main><footer>Ясноленд · Где английский становится понятным</footer>`, true);
}

function giftCard(gift, balance) {
  const price = num(get(gift,'Price')); const giftName = get(gift,'Gift Name','Gift');
  const available = norm(giftName).toLowerCase() === 'big prize' || norm(get(gift,'Available')).toLowerCase() === 'yes';
  const localGiftImages = { stickers:'gift-stickers.jpg', pencil:'gift-pencil.jpg', notebook:'gift-notebook.jpg', 'mystery box':'gift-mystery-box.jpg', 'big prize':'gift-big-prize.jpg' };
  const localGiftImage = localGiftImages[norm(giftName).toLowerCase()];
  const imageUrl = localGiftImage ? `${import.meta.env.BASE_URL}images/${localGiftImage}` : norm(get(gift,'Image URL')) || `${import.meta.env.BASE_URL}images/${encodeURIComponent(norm(get(gift,'Gift ID')))}.jpg`;
  const shortfall = price - balance; const state = !available ? 'not-available' : balance >= price ? 'available' : 'saving'; const label = !available ? 'Not available' : balance >= price ? 'Available' : `${shortfall} ${shortfall === 1 ? 'Dollar' : 'Dollars'} to go`;
  return `<article class="gift-card ${!available ? 'unavailable' : ''}"><div class="gift-image"><img src="${esc(imageUrl)}" alt="${esc(giftName)}" onerror="this.hidden=true;this.nextElementSibling.hidden=false"><div class="placeholder" hidden><svg viewBox="0 0 24 24"><path d="M20 12v9H4v-9M2 7h20v5H2zM12 7v14M12 7H7.5A2.5 2.5 0 1 1 12 4.5V7Zm0 0h4.5A2.5 2.5 0 1 0 12 4.5V7Z"/></svg></div><span class="gift-status ${state}">${label}</span></div><div class="gift-info"><small>${esc(get(gift,'Category'))}</small><h3>${esc(giftName)}</h3><b>${price} <span>Dollars</span></b></div></article>`;
}

function renderError(detail) { shell(`<main class="error-page"><div><p class="eyebrow">CONNECTION PROBLEM</p><h1>Unable to load data</h1><p>The live Excel file and the local fallback could not be opened. Check your connection and try again.</p><button class="primary" data-action="refresh">Try again</button><small>${esc(detail)}</small></div></main>`); }
function bindActions() { document.querySelector('.language-picker select')?.addEventListener('change', event => { uiLanguage = event.target.value; try { localStorage.setItem(LANGUAGE_KEY, uiLanguage); } catch {} currentView === 'progress' ? renderProgress() : renderDashboard(); }); document.querySelectorAll('[data-action="refresh"]').forEach(b => b.onclick = () => loadData(true)); document.querySelectorAll('[data-action="logout"]').forEach(b => b.onclick = () => { forgetStudent(); currentStudent = null; renderLogin(); }); document.querySelectorAll('[data-action="login"]').forEach(b => b.onclick = () => { currentStudent = null; renderLogin(); }); document.querySelectorAll('[data-action="leaderboard"]').forEach(b => b.onclick = renderLeaderboard); document.querySelectorAll('[data-tab="rewards"]').forEach(b => b.onclick = renderDashboard); document.querySelectorAll('[data-tab="progress"]').forEach(b => b.onclick = renderProgress); document.querySelectorAll('[data-action="star-rules"]').forEach(b => b.onclick = () => { const rules = document.querySelector('.star-rules'); if (rules) rules.hidden = !rules.hidden; }); document.querySelectorAll('.lesson-toggle').forEach(button => button.onclick = () => { const target = document.getElementById(button.dataset.lessons); const opening = target?.hidden; document.querySelectorAll('.lesson-list').forEach(list => list.hidden = true); document.querySelectorAll('.lesson-toggle').forEach(item => item.setAttribute('aria-expanded','false')); if (target && opening) { target.hidden = false; button.setAttribute('aria-expanded','true'); } }); }

document.addEventListener('click', event => {
  if (event.target.closest('[data-action="history"]')) document.querySelector('.dollars-history')?.scrollIntoView({ behavior:'smooth', block:'start' });
  const monthButton = event.target.closest('[data-mobile-month]');
  if (monthButton) {
    document.querySelectorAll('[data-mobile-month]').forEach(button => button.classList.toggle('active', button === monthButton));
    document.querySelectorAll('[data-month-card]').forEach(card => card.classList.toggle('active', card.dataset.monthCard === monthButton.dataset.mobileMonth));
    updateMonthlySummary();
  }
});

shell('<main class="loading-page"><div class="loader"></div><p>Loading data...</p></main>');
loadData();
