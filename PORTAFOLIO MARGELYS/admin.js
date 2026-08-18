const supabase = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

const loginBox = document.getElementById('login-box');
const panel = document.getElementById('panel');
const loginError = document.getElementById('login-error');
const formError = document.getElementById('form-error');

async function checkSession() {
  const { data: { session } } = await supabase.auth.getSession();
  if (session) {
    loginBox.style.display = 'none';
    panel.style.display = 'block';
    loadProjectList();
  } else {
    loginBox.style.display = 'block';
    panel.style.display = 'none';
  }
}

document.getElementById('login-btn').addEventListener('click', async () => {
  loginError.style.display = 'none';
  const email = document.getElementById('email').value.trim();
  const password = document.getElementById('password').value;

  if (!email || !password) {
    loginError.textContent = 'Completa email y contraseña.';
    loginError.style.display = 'block';
    return;
  }

  const { error } = await supabase.auth.signInWithPassword({ email, password });
  if (error) {
    loginError.textContent = 'No se pudo iniciar sesión. Revisa los datos.';
    loginError.style.display = 'block';
    return;
  }
  checkSession();
});

document.getElementById('logout-btn').addEventListener('click', async () => {
  await supabase.auth.signOut();
  checkSession();
});

document.getElementById('save-btn').addEventListener('click', async () => {
  formError.style.display = 'none';
  const title = document.getElementById('p-title').value.trim();
  const category = document.getElementById('p-category').value.trim();
  const cell_ref = document.getElementById('p-cellref').value.trim();
  const image_url = document.getElementById('p-image').value.trim() || null;
  const problem = document.getElementById('p-problem').value.trim();
  const result = document.getElementById('p-result').value.trim();
  const tagsRaw = document.getElementById('p-tags').value.trim();
  const featured = document.getElementById('p-featured').checked;

  if (!title || !category || !problem || !result) {
    formError.textContent = 'Completa título, categoría, problema y resultado.';
    formError.style.display = 'block';
    return;
  }

  const tags = tagsRaw ? tagsRaw.split(',').map(t => t.trim()).filter(Boolean) : [];

  const { error } = await supabase.from('projects').insert({
    title, category, cell_ref, image_url, problem, result, tags, featured, sort_order: 99
  });

  if (error) {
    formError.textContent = 'No se pudo guardar. Intenta de nuevo.';
    formError.style.display = 'block';
    return;
  }

  ['p-title', 'p-category', 'p-cellref', 'p-image', 'p-problem', 'p-result', 'p-tags'].forEach(id => document.getElementById(id).value = '');
  document.getElementById('p-featured').checked = false;
  loadProjectList();
});

async function loadProjectList() {
  const list = document.getElementById('project-list');
  const { data, error } = await supabase.from('projects').select('*').order('sort_order');
  if (error || !data) {
    list.innerHTML = '<p style="opacity:0.6; font-size:13px;">No se pudo cargar la lista.</p>';
    return;
  }
  list.innerHTML = data.map(p => `
    <div class="project-row">
      <span>${escapeHtml(p.title)}</span>
      <button data-id="${p.id}" class="delete-btn">Eliminar</button>
    </div>
  `).join('');

  list.querySelectorAll('.delete-btn').forEach(btn => {
    btn.addEventListener('click', async () => {
      await supabase.from('projects').delete().eq('id', btn.dataset.id);
      loadProjectList();
    });
  });
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str ?? '';
  return div.innerHTML;
}

checkSession();
