const supabase = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

// --- taxonomía de habilidades: se completa a mano una sola vez por herramienta nueva.
// Cualquier tag que NO esté aquí se agrega solo como burbuja nueva de primer nivel (periwinkle).
const SKILL_FAMILY = {
  'Excel Avanzado': 'gold', 'VBA': 'gold', 'Power Query': 'gold', 'Power Pivot': 'gold',
  'Power BI': 'gold', 'Dashboards': 'gold', 'Tablas Dinámicas': 'gold', 'LAMBDA': 'gold',
  'Formato Condicional': 'gold', 'Mapas': 'gold', 'Segmentación de datos': 'gold',
  'AppSheets': 'periwinkle', 'SQL': 'periwinkle', 'Python': 'periwinkle',
  'Comunicación': 'fuchsia', 'Oratoria': 'fuchsia', 'Automatización': 'fuchsia'
};
const SKILL_PARENT = {
  'Formato Condicional': 'Power Query',
  'Tablas Dinámicas': 'Excel Avanzado',
  'LAMBDA': 'Excel Avanzado',
  'Mapas': 'Power BI',
  'Segmentación de datos': 'Dashboards'
};
const FAMILY_GRAD = { gold: 'gGold', periwinkle: 'gPeri', fuchsia: 'gFuc' };

async function loadProjects() {
  const grid = document.getElementById('proyectos-grid');
  const { data, error } = await supabase.from('projects').select('*').order('sort_order');

  if (error || !data) {
    grid.innerHTML = '<p style="opacity:0.6; font-size:14px;">No se pudieron cargar los proyectos.</p>';
    console.error(error);
    return;
  }

  grid.innerHTML = data.map(renderCard).join('');
  renderSkillsChart(data);
}

function renderCard(p) {
  const tags = (p.tags || []).map(t => `<span class="tag">${esc(t)}</span>`).join('');
  const thumb = p.image_url
    ? `<img class="thumb" src="${esc(p.image_url)}" alt="${esc(p.title)}" />`
    : `<div class="thumb"></div>`;
  const catClass = p.category && p.category.length % 2 === 0 ? 'c-peri' : 'c-fuc';
  return `
    <a class="card" href="proyecto.html?id=${p.id}">
      ${thumb}
      <div class="body">
        ${p.featured ? '<span class="badge-gold badge-featured">Featured</span>' : ''}
        <p class="ref">${esc(p.cell_ref || '')}${p.cell_ref ? ' · ' : ''}${esc(p.category)}</p>
        <span class="category ${catClass}">${esc(p.category)}</span>
        <h3>${esc(p.title)}</h3>
        <p class="desc">${esc(p.result)}</p>
        <div class="tags">${tags}</div>
      </div>
    </a>
  `;
}

function renderSkillsChart(projects) {
  const counts = {};
  projects.forEach(p => (p.tags || []).forEach(t => { counts[t] = (counts[t] || 0) + 1; }));

  const allTags = Object.keys(counts);
  const children = allTags.filter(t => SKILL_PARENT[t] && allTags.includes(SKILL_PARENT[t]));
  const primary = allTags.filter(t => !children.includes(t));

  if (primary.length === 0) {
    document.getElementById('skills-chart').innerHTML = '';
    return;
  }

  // el primario con más peso propio + el de sus hijos va al centro
  const weight = t => counts[t] + children.filter(c => SKILL_PARENT[c] === t).reduce((s, c) => s + counts[c], 0);
  primary.sort((a, b) => weight(b) - weight(a));
  const centerTag = primary[0];
  const orbit = primary.slice(1);

  const W = 700, H = 420, cx = W / 2, cy = 200;
  const R = Math.min(220, W / 2 - 60);
  const rBubble = t => Math.min(58, 20 + weight(t) * 7);
  const family = t => SKILL_FAMILY[t] || 'periwinkle';

  let lines = '', circles = '', labels = '';

  const cR = rBubble(centerTag);
  circles += `<circle cx="${cx}" cy="${cy}" r="${cR}" fill="url(#${FAMILY_GRAD[family(centerTag)]})"/>`;
  labels += centerLabel(centerTag, cx, cy);

  orbit.forEach((tag, i) => {
    const angle = (i / orbit.length) * 2 * Math.PI - Math.PI / 2;
    const x = cx + R * Math.cos(angle);
    const y = cy + R * Math.sin(angle) * 0.85;
    const r = rBubble(tag);
    lines += `<line x1="${cx}" y1="${cy}" x2="${x}" y2="${y}" stroke="#E2E6EF" stroke-opacity="0.16" stroke-width="1.5"/>`;
    circles += `<circle cx="${x}" cy="${y}" r="${r}" fill="url(#${FAMILY_GRAD[family(tag)]})"/>`;
    labels += `<text x="${x}" y="${y + r + 16}" font-size="12" fill="#E2E6EF" text-anchor="middle">${esc(tag)}</text>`;

    const kids = children.filter(c => SKILL_PARENT[c] === tag);
    kids.forEach((kid, j) => {
      const kAngle = angle + (j - (kids.length - 1) / 2) * 0.5;
      const kx = x + (r + 34) * Math.cos(kAngle);
      const ky = y + (r + 34) * Math.sin(kAngle);
      const kr = Math.min(26, 12 + counts[kid] * 5);
      lines += `<line x1="${x}" y1="${y}" x2="${kx}" y2="${ky}" stroke="#E2E6EF" stroke-opacity="0.12" stroke-width="1"/>`;
      circles += `<circle cx="${kx}" cy="${ky}" r="${kr}" fill="url(#${FAMILY_GRAD[family(kid)]})"/>`;
      labels += `<text x="${kx}" y="${ky + kr + 14}" font-size="10.5" fill="#E2E6EF" text-anchor="middle" opacity="0.85">${esc(kid)}</text>`;
    });
  });

  const svg = `
  <svg viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg" style="width:100%; display:block;">
    <defs>
      <radialGradient id="gGold" cx="38%" cy="32%" r="75%"><stop offset="0%" stop-color="#F5DFAE"/><stop offset="100%" stop-color="#DFA84F"/></radialGradient>
      <radialGradient id="gPeri" cx="38%" cy="32%" r="75%"><stop offset="0%" stop-color="#ABB0EA"/><stop offset="100%" stop-color="#6E74C9"/></radialGradient>
      <radialGradient id="gFuc" cx="38%" cy="32%" r="75%"><stop offset="0%" stop-color="#C868A0"/><stop offset="100%" stop-color="#8F3062"/></radialGradient>
      <filter id="glow" x="-60%" y="-60%" width="220%" height="220%">
        <feGaussianBlur stdDeviation="6" result="b"/><feMerge><feMergeNode in="b"/><feMergeNode in="SourceGraphic"/></feMerge>
      </filter>
    </defs>
    <g stroke-linecap="round">${lines}</g>
    <g filter="url(#glow)">${circles}</g>
    ${labels}
  </svg>`;

  document.getElementById('skills-chart').innerHTML = svg;
}

function centerLabel(tag, x, y) {
  const words = tag.split(' ');
  if (words.length === 1) {
    return `<text x="${x}" y="${y + 5}" font-size="15" font-weight="600" fill="#1E2540" text-anchor="middle">${esc(tag)}</text>`;
  }
  const mid = Math.ceil(words.length / 2);
  const l1 = words.slice(0, mid).join(' '), l2 = words.slice(mid).join(' ');
  return `
    <text x="${x}" y="${y - 4}" font-size="15" font-weight="600" fill="#1E2540" text-anchor="middle">${esc(l1)}</text>
    <text x="${x}" y="${y + 16}" font-size="15" font-weight="600" fill="#1E2540" text-anchor="middle">${esc(l2)}</text>`;
}

function esc(str) {
  const div = document.createElement('div');
  div.textContent = str ?? '';
  return div.innerHTML;
}

function setupCredentials() {
  const extra = document.getElementById('cred-extra');
  document.getElementById('cred-toggle-2')?.addEventListener('click', () => extra.classList.toggle('open'));
}

setupCredentials();
loadProjects();
