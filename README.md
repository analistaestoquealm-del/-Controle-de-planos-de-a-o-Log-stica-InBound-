# -Controle-de-planos-de-a-o-Log-stica-InBound-
App (Mobile e Descktop), para controle dos planos de ação elaborados pelo setor de Logística InBound da Brametal
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Brametal - Logística Inbound</title>
    <script src="https://cdn.jsdelivr.net/npm/@tailwindcss/browser@4"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://unpkg.com/lucide@latest"></script>
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.29/jspdf.plugin.autotable.min.js"></script>
    <style>
        :root {
            --brand-red: #E30613;
            --brand-dark-red: #B8050F;
            --brand-gray: #2D3134;
        }
        .bg-brand-red { background-color: var(--brand-red); }
        .text-brand-red { color: var(--brand-red); }
        .hover\:bg-brand-dark-red:hover { background-color: var(--brand-dark-red); }
        .border-brand-red { border-color: var(--brand-red); }
        .text-brand-gray { color: var(--brand-gray); }
        .bg-brand-gray { background-color: var(--brand-gray); }
    </style>
</head>
<body class="bg-gray-50 font-sans text-gray-800 antialiased selection:bg-red-200">

    <div id="toast-container" class="fixed top-4 right-4 z-50 flex flex-col gap-2"></div>

    <div class="flex h-screen overflow-hidden">
        
        <aside id="sidebar" class="fixed inset-y-0 left-0 z-40 w-64 bg-white border-r border-gray-200 transform -translate-x-full transition-transform duration-300 ease-in-out md:translate-x-0 md:static md:flex md:flex-col shadow-sm">
            <div class="flex items-center justify-between h-16 px-4 border-b border-gray-200 bg-white">
                <div class="flex items-center gap-3">
                    <div class="w-8 h-8 rounded-full bg-brand-red relative flex items-center justify-center overflow-hidden shadow-sm animate-pulse">
                        <div class="absolute w-full h-0.5 bg-white transform rotate-45"></div>
                        <div class="absolute w-full h-0.5 bg-white transform -rotate-45"></div>
                        <div class="absolute h-full w-0.5 bg-white left-1/3"></div>
                        <div class="absolute h-full w-0.5 bg-white right-1/3"></div>
                    </div>
                    <span class="font-black tracking-tighter text-lg text-brand-gray">BRAMETAL</span>
                </div>
                <button onclick="toggleSidebar()" class="md:hidden text-gray-400 hover:text-brand-red transition-colors duration-200">
                    <i data-lucide="x" class="w-6 h-6"></i>
                </button>
            </div>
            
            <nav class="flex-1 p-4 space-y-1.5 overflow-y-auto">
                <button onclick="switchScreen('create-project-screen')" class="nav-btn w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-bold transition-all duration-300 transform hover:translate-x-1 hover:bg-red-50 text-gray-700" id="btn-create-project-screen">
                    <i data-lucide="plus-circle" class="w-5 h-5 text-brand-red transition-transform duration-300 group-hover:rotate-90"></i> Criar Novo Projeto
                </button>
                <button onclick="switchScreen('ongoing-projects-screen')" class="nav-btn w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-bold transition-all duration-300 transform hover:translate-x-1 hover:bg-gray-100 text-gray-600" id="btn-ongoing-projects-screen">
                    <i data-lucide="refresh-cw" class="w-5 h-5 text-brand-gray"></i> Projetos em Andamento
                </button>
                <button onclick="switchScreen('completed-projects-screen')" class="nav-btn w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-bold transition-all duration-300 transform hover:translate-x-1 hover:bg-gray-100 text-gray-600" id="btn-completed-projects-screen">
                    <i data-lucide="check-circle" class="w-5 h-5 text-brand-gray"></i> Projetos Concluídos
                </button>
                <div class="h-px bg-gray-200 my-4"></div>
                <button onclick="switchScreen('dashboard-screen')" class="nav-btn w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-bold transition-all duration-300 transform hover:translate-x-1 hover:bg-gray-100 text-gray-600" id="btn-dashboard-screen">
                    <i data-lucide="bar-chart-3" class="w-5 h-5 text-brand-red"></i> Dashboard Indicadores
                </button>
            </nav>
        </aside>

        <div class="flex-1 flex flex-col min-w-0 overflow-hidden">
            
            <header class="h-16 bg-white border-b border-gray-200 flex items-center justify-between px-4 md:px-6 z-30 shadow-xs">
                <button onclick="toggleSidebar()" class="md:hidden p-2 rounded-md text-gray-600 hover:bg-gray-100 transition-colors duration-200">
                    <i data-lucide="menu" class="w-6 h-6"></i>
                </button>
                <h1 id="screen-title" class="text-xl font-black text-brand-gray truncate">Criar Novo Projeto</h1>
                <div class="flex items-center gap-3">
                    <span class="text-xs font-bold bg-gray-100 border border-gray-200 px-3 py-1 rounded-full text-brand-gray hidden sm:inline-block">Inbound Database Control v2.0</span>
                </div>
            </header>

            <main class="flex-1 overflow-y-auto p-4 md:p-6 bg-gray-50">

                <section id="create-project-screen" class="screen-content space-y-6">
                    <div class="bg-white rounded-xl shadow-xs border border-gray-200 p-4 md:p-6">
                        <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
                            <div>
                                <label class="block text-xs font-bold uppercase tracking-wider text-gray-400 mb-1">ID do Projeto</label>
                                <input type="text" id="proj-id" readonly class="w-full bg-gray-50 border border-gray-200 rounded-lg px-3 py-2 text-sm font-mono focus:outline-none font-bold text-brand-gray" value="PROJ-0001">
                            </div>
                            <div>
                                <label class="block text-xs font-bold uppercase tracking-wider text-gray-500 mb-1">Nome do Projeto *</label>
                                <input type="text" id="proj-name" class="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-red-100 focus:border-brand-red focus:outline-none transition-all duration-200" placeholder="Ex: Otimização de Recebimento de Aço">
                            </div>
                            <div>
                                <label class="block text-xs font-bold uppercase tracking-wider text-gray-500 mb-1">Categoria *</label>
                                <select id="proj-category" class="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-red-100 focus:border-brand-red focus:outline-none transition-all duration-200">
                                    <option value="Melhorias de Sistemas">Melhorias de Sistemas</option>
                                    <option value="Melhorias de Processos">Melhorias de Processos</option>
                                    <option value="Estratégico">Estratégico</option>
                                    <option value="Operacional">Operacional</option>
                                    <option value="Corretivas">Corretivas</option>
                                    <option value="Preventivas">Preventivas</option>
                                </select>
                            </div>
                            <div>
                                <label class="block text-xs font-bold uppercase tracking-wider text-gray-500 mb-1">Criador do Projeto *</label>
                                <input type="text" id="proj-creator" class="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-red-100 focus:border-brand-red focus:outline-none transition-all duration-200" placeholder="Nome do Responsável">
                            </div>
                            <div>
                                <label class="block text-xs font-bold uppercase tracking-wider text-gray-500 mb-1">Data de Criação *</label>
                                <input type="date" id="proj-date" class="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-red-100 focus:border-brand-red focus:outline-none transition-all duration-200">
                            </div>
                        </div>
                    </div>

                    <div class="bg-white rounded-xl shadow-xs border border-gray-200 overflow-hidden flex flex-col">
                        <div class="p-4 bg-gray-50 border-b border-gray-200 flex flex-wrap gap-3 items-center justify-between">
                            <h3 class="font-bold text-brand-gray flex items-center gap-2">
                                <i data-lucide="list-todo" class="w-5 h-5 text-brand-red"></i> Plano de Ações
                            </h3>
                            <div class="flex flex-wrap gap-2">
                                <button onclick="addGridRow()" class="flex items-center gap-1.5 bg-brand-red hover:bg-brand-dark-red text-white text-xs font-bold px-3 py-2 rounded-lg transition-all duration-200 shadow-xs active:scale-95 cursor-pointer">
                                    <i data-lucide="plus" class="w-4 h-4"></i> Adicionar Linha
                                </button>
                                <button onclick="deleteSelectedRows()" class="flex items-center gap-1.5 bg-gray-200 hover:bg-gray-300 text-gray-700 text-xs font-bold px-3 py-2 rounded-lg transition-all duration-200 active:scale-95 cursor-pointer">
                                    <i data-lucide="trash-2" class="w-4 h-4 text-brand-red"></i> Excluir Selecionados
                                </button>
                                <button onclick="exportActions('excel')" class="flex items-center gap-1.5 bg-gray-600 hover:bg-brand-gray text-white text-xs font-bold px-3 py-2 rounded-lg transition-all duration-200 shadow-xs active:scale-95 cursor-pointer">
                                    <i data-lucide="file-spreadsheet" class="w-4 h-4"></i> Excel
                                </button>
                                <button onclick="exportActions('pdf')" class="flex items-center gap-1.5 bg-gray-700 hover:bg-gray-800 text-white text-xs font-bold px-3 py-2 rounded-lg transition-all duration-200 shadow-xs active:scale-95 cursor-pointer">
                                    <i data-lucide="file-text" class="w-4 h-4"></i> PDF
                                </button>
                            </div>
                        </div>

                        <div class="overflow-x-auto">
                            <table class="w-full text-left text-xs border-collapse min-w-[1200px]" id="actions-table">
                                <thead class="bg-gray-100 text-gray-600 uppercase font-bold tracking-wider border-b border-gray-200">
                                    <tr>
                                        <th class="p-3 w-10 text-center"><input type="checkbox" id="select-all-actions" onchange="toggleSelectAllActions(this)"></th>
                                        <th class="p-3 w-36">Data Criação *</th>
                                        <th class="p-3 w-64">Descrição da Ação *</th>
                                        <th class="p-3 w-44">Responsável *</th>
                                        <th class="p-3 w-36">Início Planejado *</th>
                                        <th class="p-3 w-36">Término Estimado *</th>
                                        <th class="p-3 w-36">Término Real</th>
                                        <th class="p-3 w-40">Status *</th>
                                        <th class="p-3 w-28 text-center">Status KPI</th>
                                        <th class="p-3 w-48">Observações</th>
                                        <th class="p-3 w-12 text-center">Ações</th>
                                    </tr>
                                </thead>
                                <tbody id="actions-grid-body" class="divide-y divide-gray-200">
                                    </tbody>
                            </table>
                        </div>
                        
                        <div class="p-4 border-t border-gray-200 bg-gray-50 flex justify-end">
                            <button onclick="saveCurrentProject()" class="flex items-center gap-2 bg-brand-red hover:bg-brand-dark-red text-white font-black text-sm px-6 py-2.5 rounded-lg shadow-sm transition-all duration-300 active:scale-95 cursor-pointer">
                                <i data-lucide="save" class="w-4 h-4"></i> Salvar Projeto no Banco
                            </button>
                        </div>
                    </div>
                </section>

                <section id="projects-list-screen" class="screen-content space-y-4 hidden">
                    <div class="bg-white rounded-xl shadow-xs border border-gray-200 p-4 flex flex-wrap gap-3 items-center justify-between">
                        <div class="relative flex-1 max-w-md">
                            <i data-lucide="search" class="w-4 h-4 absolute left-3 top-3 text-gray-400"></i>
                            <input type="text" id="search-projects" oninput="renderProjectLists()" class="w-full border border-gray-300 rounded-lg pl-9 pr-4 py-2 text-sm focus:outline-none focus:ring-1 focus:ring-brand-red transition-all" placeholder="Buscar por ID ou Nome do Projeto...">
                        </div>
                        <div class="flex gap-2">
                            <button onclick="deleteMultipleProjects()" class="flex items-center gap-1.5 bg-brand-red hover:bg-brand-dark-red text-white text-xs font-bold px-3 py-2 rounded-lg transition-all shadow-xs active:scale-95 cursor-pointer">
                                <i data-lucide="trash-2" class="w-4 h-4"></i> Excluir Selecionados
                            </button>
                            <button onclick="exportProjects('excel')" class="flex items-center gap-1.5 bg-gray-600 hover:bg-brand-gray text-white text-xs font-bold px-3 py-2 rounded-lg transition-all shadow-xs cursor-pointer">
                                <i data-lucide="file-spreadsheet" class="w-4 h-4"></i> Exportar Lista Excel
                            </button>
                        </div>
                    </div>

                    <div class="bg-white rounded-xl shadow-xs border border-gray-200 overflow-hidden">
                        <div class="overflow-x-auto">
                            <table class="w-full text-left text-sm border-collapse">
                                <thead class="bg-gray-100 text-gray-600 uppercase font-bold text-xs tracking-wider border-b border-gray-200">
                                    <tr>
                                        <th class="p-4 w-12 text-center"><input type="checkbox" id="select-all-projects" onchange="toggleSelectAllProjects(this)"></th>
                                        <th class="p-4 w-28">ID Projeto</th>
                                        <th class="p-4">Nome do Projeto</th>
                                        <th class="p-4 w-44">Categoria</th>
                                        <th class="p-4 w-44">Criador</th>
                                        <th class="p-4 w-36">Criação</th>
                                        <th class="p-4 w-32 text-center">Ações Concluídas</th>
                                        <th class="p-4 w-40 text-center">Ações</th>
                                    </tr>
                                </thead>
                                <tbody id="project-list-body" class="divide-y divide-gray-200">
                                    </tbody>
                            </table>
                        </div>
                    </div>
                </section>

                <section id="dashboard-screen" class="screen-content space-y-6 hidden">
                    <div class="bg-white rounded-xl shadow-xs border border-gray-200 p-4 grid grid-cols-1 sm:grid-cols-4 gap-4 items-end">
                        <div>
                            <label class="block text-xs font-bold uppercase tracking-wider text-gray-500 mb-1">Data Início</label>
                            <input type="date" id="filter-start-date" onchange="updateDashboard()" class="w-full border border-gray-300 rounded-lg px-3 py-2 text-xs focus:outline-none focus:border-brand-red transition-all">
                        </div>
                        <div>
                            <label class="block text-xs font-bold uppercase tracking-wider text-gray-500 mb-1">Data Fim</label>
                            <input type="date" id="filter-end-date" onchange="updateDashboard()" class="w-full border border-gray-300 rounded-lg px-3 py-2 text-xs focus:outline-none focus:border-brand-red transition-all">
                        </div>
                        <div>
                            <label class="block text-xs font-bold uppercase tracking-wider text-gray-500 mb-1">Filtrar ID Projeto</label>
                            <select id="filter-proj-id" onchange="updateDashboard()" class="w-full border border-gray-300 rounded-lg px-3 py-2 text-xs focus:outline-none focus:border-brand-red transition-all">
                                <option value="">Todos</option>
                            </select>
                        </div>
                        <div>
                            <label class="block text-xs font-bold uppercase tracking-wider text-gray-500 mb-1">Filtrar Categoria</label>
                            <select id="filter-category" onchange="updateDashboard()" class="w-full border border-gray-300 rounded-lg px-3 py-2 text-xs focus:outline-none focus:border-brand-red transition-all">
                                <option value="">Todas</option>
                                <option value="Melhorias de Sistemas">Melhorias de Sistemas</option>
                                <option value="Melhorias de Processos">Melhorias de Processos</option>
                                <option value="Estratégico">Estratégico</option>
                                <option value="Operacional">Operacional</option>
                                <option value="Corretivas">Corretivas</option>
                                <option value="Preventivas">Preventivas</option>
                            </select>
                        </div>
                    </div>

                    <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                        <div class="bg-white p-5 border border-gray-200 rounded-xl shadow-xs flex items-center justify-between transition-all duration-300 hover:shadow-md">
                            <div>
                                <p class="text-xs font-bold text-gray-400 uppercase tracking-wider">Total de Projetos</p>
                                <h3 id="dash-total-projects" class="text-2xl font-black text-brand-gray mt-1">0</h3>
                            </div>
                            <div class="p-3 bg-red-50 rounded-lg text-brand-red"><i data-lucide="folder" class="w-6 h-6"></i></div>
                        </div>
                        <div class="bg-white p-5 border border-gray-200 rounded-xl shadow-xs flex items-center justify-between transition-all duration-300 hover:shadow-md">
                            <div>
                                <p class="text-xs font-bold text-gray-400 uppercase tracking-wider">Total de Ações</p>
                                <h3 id="dash-total-actions" class="text-2xl font-black text-brand-gray mt-1">0</h3>
                            </div>
                            <div class="p-3 bg-gray-100 rounded-lg text-brand-gray"><i data-lucide="activity" class="w-6 h-6"></i></div>
                        </div>
                        <div class="bg-white p-5 border border-gray-200 rounded-xl shadow-xs flex items-center justify-between transition-all duration-300 hover:shadow-md">
                            <div>
                                <p class="text-xs font-bold text-gray-400 uppercase tracking-wider">Ações Atrasadas</p>
                                <h3 id="dash-delayed-actions" class="text-2xl font-black text-brand-red mt-1">0</h3>
                            </div>
                            <div class="p-3 bg-red-100 rounded-lg text-brand-red"><i data-lucide="alert-triangle" class="w-6 h-6"></i></div>
                        </div>
                        <div class="bg-white p-5 border border-gray-200 rounded-xl shadow-xs flex items-center justify-between transition-all duration-300 hover:shadow-md">
                            <div>
                                <p class="text-xs font-bold text-gray-400 uppercase tracking-wider">Taxa de Conclusão</p>
                                <h3 id="dash-completion-rate" class="text-2xl font-black text-brand-gray mt-1">0%</h3>
                            </div>
                            <div class="p-3 bg-gray-50 rounded-lg text-gray-500"><i data-lucide="trending-up" class="w-6 h-6"></i></div>
                        </div>
                    </div>

                    <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                        <div class="bg-white p-4 md:p-6 border border-gray-200 rounded-xl shadow-xs flex flex-col items-center">
                            <h4 class="text-sm font-black text-brand-gray self-start mb-4">Total de Projetos por Categoria</h4>
                            <div class="w-full h-64"><canvas id="chart-projects-category"></canvas></div>
                        </div>
                        <div class="bg-white p-4 md:p-6 border border-gray-200 rounded-xl shadow-xs flex flex-col items-center">
                            <h4 class="text-sm font-black text-brand-gray self-start mb-4">Ações por Status</h4>
                            <div class="w-full h-64"><canvas id="chart-actions-status"></canvas></div>
                        </div>
                        <div class="bg-white p-4 md:p-6 border border-gray-200 rounded-xl shadow-xs flex flex-col items-center">
                            <h4 class="text-sm font-black text-brand-gray self-start mb-4">Ações por KPI de Prazo</h4>
                            <div class="w-full h-64"><canvas id="chart-actions-kpi"></canvas></div>
                        </div>
                        <div class="bg-white p-4 md:p-6 border border-gray-200 rounded-xl shadow-xs flex flex-col items-center">
                            <h4 class="text-sm font-black text-brand-gray self-start mb-4">Volume de Ações por Projeto</h4>
                            <div class="w-full h-64"><canvas id="chart-actions-project"></canvas></div>
                        </div>
                    </div>
                </section>

            </main>
        </div>
    </div>

    <script>
        // Application Global State
        let state = {
            currentScreen: 'create-project-screen',
            projects: [],
            currentEditingProject: null,
            listFilterMode: 'ongoing'
        };

        // Identidade Visual baseada na logo da Brametal
        const PALETTE = {
            red: '#E30613',
            darkRed: '#B8050F',
            brandGray: '#2D3134',
            mediumGray: '#6B7280',
            lightGray: '#D1D5DB',
            ultraLightGray: '#F3F4F6',
            white: '#FFFFFF'
        };

        // Banco de Dados Executivo Local (IndexedDB) para Histórico e Consultas
        let db;
        const DB_NAME = 'BrametalInboundDB';
        const DB_VERSION = 1;

        function initDatabase() {
            return new Promise((resolve, reject) => {
                const request = indexedDB.open(DB_NAME, DB_VERSION);
                
                request.onerror = (event) => {
                    showToast("Erro ao abrir banco de dados histórico.", "error");
                    reject(event);
                };

                request.onsuccess = (event) => {
                    db = event.target.result;
                    loadProjectsFromDB().then(resolve);
                };

                request.onupgradeneeded = (event) => {
                    const database = event.target.result;
                    if (!database.objectStoreNames.contains('projects')) {
                        database.createObjectStore('projects', { keyPath: 'id' });
                    }
                };
            });
        }

        function loadProjectsFromDB() {
            return new Promise((resolve) => {
                const transaction = db.transaction(['projects'], 'readonly');
                const store = transaction.objectStore('projects');
                const request = store.getAll();

                request.onsuccess = (event) => {
                    state.projects = event.target.result || [];
                    resolve();
                };
            });
        }

        function syncProjectToDB(project) {
            return new Promise((resolve) => {
                const transaction = db.transaction(['projects'], 'readwrite');
                const store = transaction.objectStore('projects');
                store.put(project);
                transaction.oncomplete = () => resolve();
            });
        }

        function deleteProjectFromDB(id) {
            return new Promise((resolve) => {
                const transaction = db.transaction(['projects'], 'readwrite');
                const store = transaction.objectStore('projects');
                store.delete(id);
                transaction.oncomplete = () => resolve();
            });
        }

        // Inicialização do Sistema
        document.addEventListener('DOMContentLoaded', () => {
            initDatabase().then(() => {
                lucide.createIcons();
                document.getElementById('proj-date').value = new Date().toISOString().split('T')[0];
                generateNewProjectId();
                addGridRow();
                updateDashboardDropdowns();
            });
        });

        // Alertas Flutuantes (Toast)
        function showToast(message, type = 'success') {
            const container = document.getElementById('toast-container');
            const toast = document.createElement('div');
            const colorClass = type === 'success' ? 'bg-brand-gray text-white border-l-4 border-brand-red' : 'bg-brand-gray text-red-300 border-l-4 border-red-500';
            toast.className = `flex items-center gap-3 px-4 py-3 rounded-lg shadow-xl text-xs font-bold ${colorClass} min-w-[300px] transform translate-y-2 opacity-0 transition-all duration-300`;
            toast.innerHTML = `<i data-lucide="${type === 'success' ? 'check-circle' : 'alert-circle'}" class="w-4 h-4"></i><span>${message}</span>`;
            container.appendChild(toast);
            lucide.createIcons({attrs: {class: 'w-4 h-4'}});
            setTimeout(() => { toast.classList.remove('translate-y-2', 'opacity-0'); }, 10);
            setTimeout(() => {
                toast.classList.add('opacity-0', 'translate-y-2');
                setTimeout(() => toast.remove(), 300);
            }, 4000);
        }

        function toggleSidebar() {
            const sidebar = document.getElementById('sidebar');
            sidebar.classList.toggle('-translate-x-full');
        }

        function switchScreen(screenId) {
            document.querySelectorAll('.screen-content').forEach(s => s.classList.add('hidden'));
            document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('bg-red-50', 'text-brand-red', 'bg-gray-100', 'text-gray-700'));
            
            state.currentScreen = screenId;
            const targetScreen = document.getElementById(screenId);
            if (targetScreen) targetScreen.classList.remove('hidden');

            const titleMap = {
                'create-project-screen': state.currentEditingProject ? 'Editar Projeto' : 'Criar Novo Projeto',
                'ongoing-projects-screen': 'Projetos em Andamento',
                'completed-projects-screen': 'Projetos Concluídos',
                'dashboard-screen': 'Dashboard de Indicadores'
            };

            if(screenId === 'ongoing-projects-screen' || screenId === 'completed-projects-screen') {
                state.listFilterMode = screenId === 'ongoing-projects-screen' ? 'ongoing' : 'completed';
                document.getElementById('projects-list-screen').classList.remove('hidden');
                document.getElementById('screen-title').innerText = titleMap[screenId];
                document.getElementById(`btn-${screenId}`).classList.add('bg-gray-100', 'text-gray-700');
                renderProjectLists();
            } else {
                document.getElementById('screen-title').innerText = titleMap[screenId] || '';
                document.getElementById(`btn-${screenId}`).classList.add('bg-red-50', 'text-brand-red');
            }

            if(screenId === 'dashboard-screen') {
                updateDashboardDropdowns();
                updateDashboard();
            }

            if (window.innerWidth < 768) toggleSidebar();
        }

        function generateNewProjectId() {
            if (state.projects.length === 0) {
                document.getElementById('proj-id').value = 'PROJ-0001';
                return;
            }
            const ids = state.projects.map(p => parseInt(p.id.replace('PROJ-', '')));
            const maxId = Math.max(...ids);
            document.getElementById('proj-id').value = `PROJ-${String(maxId + 1).padStart(4, '0')}`;
        }

        // Tabela Dinâmica do Plano de Ações
        function addGridRow(actionData = null) {
            const tbody = document.getElementById('actions-grid-body');
            const rowId = 'row_' + Math.random().toString(36).substring(2, 9);
            const todayStr = new Date().toISOString().split('T')[0];

            const tr = document.createElement('tr');
            tr.id = rowId;
            tr.className = "hover:bg-gray-50/80 transition-colors duration-150";

            tr.innerHTML = `
                <td class="p-2 text-center border-b border-gray-200"><input type="checkbox" class="row-checkbox rounded accent-red-600"></td>
                <td class="p-2 border-b border-gray-200"><input type="date" class="act-created w-full bg-transparent border-0 focus:ring-1 focus:ring-brand-red rounded p-1" value="${actionData?.created || todayStr}"></td>
                <td class="p-2 border-b border-gray-200"><input type="text" class="act-desc w-full bg-transparent border-0 focus:ring-1 focus:ring-brand-red rounded p-1 font-medium" placeholder="Descrição da ação" value="${actionData?.desc || ''}"></td>
                <td class="p-2 border-b border-gray-200"><input type="text" class="act-owner w-full bg-transparent border-0 focus:ring-1 focus:ring-brand-red rounded p-1" placeholder="Responsável" value="${actionData?.owner || ''}"></td>
                <td class="p-2 border-b border-gray-200"><input type="date" class="act-start w-full bg-transparent border-0 focus:ring-1 focus:ring-brand-red rounded p-1" value="${actionData?.start || todayStr}" onchange="recalculateKpi('${rowId}')"></td>
                <td class="p-2 border-b border-gray-200"><input type="date" class="act-est-end w-full bg-transparent border-0 focus:ring-1 focus:ring-brand-red rounded p-1" value="${actionData?.estEnd || todayStr}" onchange="recalculateKpi('${rowId}')"></td>
                <td class="p-2 border-b border-gray-200"><input type="date" class="act-act-end w-full bg-transparent border-0 focus:ring-1 focus:ring-brand-red rounded p-1" value="${actionData?.actEnd || ''}"></td>
                <td class="p-2 border-b border-gray-200">
                    <select class="act-status w-full bg-transparent border-0 focus:ring-1 focus:ring-brand-red rounded p-1 font-bold" onchange="handleStatusChange('${rowId}')">
                        <option value="Not Started" ${actionData?.status === 'Not Started' ? 'selected' : ''}>Not Started</option>
                        <option value="In Progress" ${actionData?.status === 'In Progress' ? 'selected' : ''}>In Progress</option>
                        <option value="Completed" ${actionData?.status === 'Completed' ? 'selected' : ''}>Completed</option>
                    </select>
                </td>
                <td class="p-2 text-center border-b border-gray-200 font-bold" id="kpi_${rowId}">-</td>
                <td class="p-2 border-b border-gray-200"><textarea class="act-notes w-full bg-transparent border-0 focus:ring-1 focus:ring-brand-red rounded p-1 resize-none h-8 text-[11px]" placeholder="Obs...">${actionData?.notes || ''}</textarea></td>
                <td class="p-2 text-center border-b border-gray-200">
                    <button onclick="deleteSingleRow('${rowId}')" class="text-gray-400 hover:text-brand-red transition-colors duration-200 transform hover:scale-110 cursor-pointer"><i data-lucide="trash-2" class="w-4 h-4"></i></button>
                </td>
            `;

            tbody.appendChild(tr);
            lucide.createIcons({attrs: {class: 'w-4 h-4'}});
            recalculateKpi(rowId);
        }

        function handleStatusChange(rowId) {
            const row = document.getElementById(rowId);
            const status = row.querySelector('.act-status').value;
            const actEndInput = row.querySelector('.act-act-end');
            
            if (status === 'Completed' && !actEndInput.value) {
                actEndInput.value = new Date().toISOString().split('T')[0];
            } else if (status !== 'Completed') {
                actEndInput.value = '';
            }
            recalculateKpi(rowId);
        }

        function recalculateKpi(rowId) {
            const row = document.getElementById(rowId);
            if (!row) return;

            const estEnd = row.querySelector('.act-est-end').value;
            const status = row.querySelector('.act-status').value;
            const kpiCell = document.getElementById(`kpi_${rowId}`);
            const today = new Date().toISOString().split('T')[0];

            if (!estEnd) {
                kpiCell.innerHTML = '-';
                return;
            }

            if (status === 'Completed') {
                kpiCell.innerHTML = `<span class="px-2 py-0.5 rounded text-[10px] bg-gray-200 text-brand-gray border border-gray-300 font-bold">Em Dia</span>`;
            } else if (today > estEnd) {
                kpiCell.innerHTML = `<span class="px-2 py-0.5 rounded text-[10px] bg-brand-red text-white font-black">Atrasado</span>`;
            } else {
                kpiCell.innerHTML = `<span class="px-2 py-0.5 rounded text-[10px] bg-brand-gray text-white font-bold">No Prazo</span>`;
            }
        }

        function toggleSelectAllActions(master) {
            document.querySelectorAll('#actions-grid-body .row-checkbox').forEach(cb => cb.checked = master.checked);
        }

        function deleteSingleRow(rowId) {
            const row = document.getElementById(rowId);
            if (row) row.remove();
        }

        function deleteSelectedRows() {
            const checkboxes = document.querySelectorAll('#actions-grid-body .row-checkbox:checked');
            if(checkboxes.length === 0) {
                showToast("Selecione ao menos uma linha para excluir.", "error");
                return;
            }
            checkboxes.forEach(cb => cb.closest('tr').remove());
            document.getElementById('select-all-actions').checked = false;
            showToast("Ações removidas da grade temporária.");
        }

        // Salvar Projeto com Persistência em Banco de Dados Real (Histórico)
        function saveCurrentProject() {
            const id = document.getElementById('proj-id').value;
            const name = document.getElementById('proj-name').value.trim();
            const category = document.getElementById('proj-category').value;
            const creator = document.getElementById('proj-creator').value.trim();
            const date = document.getElementById('proj-date').value;

            if (!name || !creator || !date) {
                showToast("Preencha todos os campos obrigatórios do cabeçalho do projeto.", "error");
                return;
            }

            const rows = document.querySelectorAll('#actions-grid-body tr');
            if (rows.length === 0) {
                showToast("Adicione pelo menos uma ação ao projeto.", "error");
                return;
            }

            let actions = [];
            let allValid = true;

            rows.forEach(row => {
                const created = row.querySelector('.act-created').value;
                const desc = row.querySelector('.act-desc').value.trim();
                const owner = row.querySelector('.act-owner').value.trim();
                const start = row.querySelector('.act-start').value;
                const estEnd = row.querySelector('.act-est-end').value;
                const actEnd = row.querySelector('.act-act-end').value;
                const status = row.querySelector('.act-status').value;
                const notes = row.querySelector('.act-notes').value.trim();

                if (!created || !desc || !owner || !start || !estEnd || !status) {
                    allValid = false;
                }

                const today = new Date().toISOString().split('T')[0];
                let kpi = "No Prazo";
                if (status === 'Completed') kpi = "Em Dia";
                else if (today > estEnd) kpi = "Atrasado";

                actions.push({ created, desc, owner, start, estEnd, actEnd, status, notes, kpi });
            });

            if (!allValid) {
                showToast("Certifique-se de que todas as linhas de ação possuem dados obrigatórios preenchidos.", "error");
                return;
            }

            const projectStatus = actions.every(a => a.status === 'Completed') ? 'Completed' : 'Ongoing';
            const projectPayload = { id, name, category, creator, date, status: projectStatus, actions };

            const index = state.projects.findIndex(p => p.id === id);
            if (index > -1) {
                state.projects[index] = projectPayload;
            } else {
                state.projects.push(projectPayload);
            }

            // Commita no Banco de Dados Histórico (IndexedDB)
            syncProjectToDB(projectPayload).then(() => {
                showToast(`Projeto gravado com sucesso no Banco de Dados! Status: ${projectStatus === 'Completed' ? 'Concluído' : 'Em Andamento'}`);
                resetProjectForm();
                switchScreen(projectStatus === 'Completed' ? 'completed-projects-screen' : 'ongoing-projects-screen');
            });
        }

        // Seleção Múltipla e Exclusão em Massa de Projetos
        function toggleSelectAllProjects(master) {
            document.querySelectorAll('#project-list-body .project-checkbox').forEach(cb => cb.checked = master.checked);
        }

        function deleteMultipleProjects() {
            const selectedCheckboxes = document.querySelectorAll('#project-list-body .project-checkbox:checked');
            if (selectedCheckboxes.length === 0) {
                showToast("Nenhum projeto selecionado para exclusão.", "error");
                return;
            }

            if (confirm(`Deseja deletar em massa os ${selectedCheckboxes.length} projetos selecionados e remover seus históricos do banco?`)) {
                const promises = Array.from(selectedCheckboxes).map(cb => {
                    const id = cb.getAttribute('data-id');
                    state.projects = state.projects.filter(p => p.id !== id);
                    return deleteProjectFromDB(id);
                });

                Promise.all(promises).then(() => {
                    showToast("Exclusão em massa concluída no banco de dados.");
                    document.getElementById('select-all-projects').checked = false;
                    renderProjectLists();
                });
            }
        }

        function deleteWholeProject(projId) {
            if(confirm(`Tem certeza de que deseja excluir permanentemente o projeto ${projId} e limpar o histórico no banco de dados?`)) {
                state.projects = state.projects.filter(p => p.id !== projId);
                deleteProjectFromDB(projId).then(() => {
                    showToast(`Projeto ${projId} apagado do histórico.`);
                    renderProjectLists();
                });
            }
        }

        function resetProjectForm() {
            state.currentEditingProject = null;
            generateNewProjectId();
            document.getElementById('proj-name').value = '';
            document.getElementById('proj-creator').value = '';
            document.getElementById('proj-date').value = new Date().toISOString().split('T')[0];
            document.getElementById('actions-grid-body').innerHTML = '';
            document.getElementById('select-all-actions').checked = false;
            addGridRow();
        }

        function loadProjectIntoForm(projId) {
            const project = state.projects.find(p => p.id === projId);
            if (!project) return;

            state.currentEditingProject = project;
            document.getElementById('proj-id').value = project.id;
            document.getElementById('proj-name').value = project.name;
            document.getElementById('proj-category').value = project.category || 'Melhorias de Sistemas';
            document.getElementById('proj-creator').value = project.creator;
            document.getElementById('proj-date').value = project.date;

            const tbody = document.getElementById('actions-grid-body');
            tbody.innerHTML = '';
            
            project.actions.forEach(act => addGridRow(act));
            switchScreen('create-project-screen');
        }

        // Renderização Dinâmica das Listas
        function renderProjectLists() {
            const tbody = document.getElementById('project-list-body');
            tbody.innerHTML = '';
            
            const searchTerm = document.getElementById('search-projects').value.toLowerCase();

            const filteredProjects = state.projects.filter(p => {
                const matchesStatus = state.listFilterMode === 'completed' ? p.status === 'Completed' : p.status === 'Ongoing';
                const matchesSearch = p.id.toLowerCase().includes(searchTerm) || p.name.toLowerCase().includes(searchTerm);
                return matchesStatus && matchesSearch;
            });

            if(filteredProjects.length === 0) {
                tbody.innerHTML = `<tr><td colspan="8" class="p-8 text-center text-gray-400 font-medium">Nenhum projeto registrado nesta visualização do histórico.</td></tr>`;
                return;
            }

            filteredProjects.forEach(p => {
                const completedCount = p.actions.filter(a => a.status === 'Completed').length;
                const totalActions = p.actions.length;

                const tr = document.createElement('tr');
                tr.className = "hover:bg-red-50/20 transition-colors duration-200 cursor-pointer border-b border-gray-100";
                tr.onclick = () => loadProjectIntoForm(p.id);

                tr.innerHTML = `
                    <td class="p-4 text-center" onclick="event.stopPropagation()"><input type="checkbox" class="project-checkbox rounded accent-red-600" data-id="${p.id}"></td>
                    <td class="p-4 font-mono font-bold text-brand-gray">${p.id}</td>
                    <td class="p-4 font-bold text-gray-800">${p.name}</td>
                    <td class="p-4 text-gray-600"><span class="bg-gray-100 px-2 py-1 rounded border border-gray-200 text-xs font-semibold">${p.category}</span></td>
                    <td class="p-4 text-gray-600 font-medium">${p.creator}</td>
                    <td class="p-4 text-gray-500">${p.date.split('-').reverse().join('/')}</td>
                    <td class="p-4 text-center">
                        <span class="font-bold text-xs bg-brand-gray text-white px-2 py-1 rounded-full">${completedCount}/${totalActions}</span>
                    </td>
                    <td class="p-4 text-center flex items-center justify-center gap-3" onclick="event.stopPropagation()">
                        <button onclick="loadProjectIntoForm('${p.id}')" class="text-gray-500 hover:text-brand-red font-bold text-xs flex items-center gap-1 transition-all duration-200 hover:scale-105 cursor-pointer">
                            <i data-lucide="edit-3" class="w-4 h-4"></i> Editar
                        </button>
                        <span class="text-gray-300">|</span>
                        <button onclick="deleteWholeProject('${p.id}')" class="text-gray-400 hover:text-brand-red font-bold text-xs flex items-center gap-1 transition-all duration-200 hover:scale-105 cursor-pointer">
                            <i data-lucide="trash-2" class="w-4 h-4"></i> Excluir
                        </button>
                    </td>
                `;
                tbody.appendChild(tr);
            });
            lucide.createIcons({attrs: {class: 'w-4 h-4'}});
        }

        // Exportação de Dados Relatórios
        function exportActions(format) {
            const rows = document.querySelectorAll('#actions-grid-body tr');
            if(rows.length === 0) {
                showToast("Não há ações carregadas para exportação.", "error");
                return;
            }

            const projId = document.getElementById('proj-id').value;
            let data = [];
            
            rows.forEach(row => {
                data.push({
                    "Criado Em": row.querySelector('.act-created').value,
                    "Descrição": row.querySelector('.act-desc').value,
                    "Responsável": row.querySelector('.act-owner').value,
                    "Início": row.querySelector('.act-start').value,
                    "Fim Estimado": row.querySelector('.act-est-end').value,
                    "Fim Real": row.querySelector('.act-act-end').value || '-',
                    "Status": row.querySelector('.act-status').value,
                    "Observações": row.querySelector('.act-notes').value
                });
            });

            if (format === 'excel') {
                const worksheet = XLSX.utils.json_to_sheet(data);
                const workbook = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(workbook, worksheet, "Ações");
                XLSX.writeFile(workbook, `Plano_de_Acoes_${projId}.xlsx`);
            } else if (format === 'pdf') {
                const { jsPDF } = window.jspdf;
                const doc = new jsPDF();
                doc.text(`Plano de Ações - ${projId}`, 14, 15);
                const tableRows = data.map(d => Object.values(d));
                const headers = Object.keys(data[0]);
                doc.autoTable({ head: [headers], body: tableRows, startY: 22, theme: 'grid', styles: { fontSize: 7 } });
                doc.save(`Plano_de_Acoes_${projId}.pdf`);
            }
        }

        function exportProjects(format) {
            if(state.projects.length === 0) {
                showToast("Não há dados históricos salvos para exportar.", "error");
                return;
            }
            const data = state.projects.map(p => ({
                "ID Projeto": p.id,
                "Nome": p.name,
                "Categoria": p.category,
                "Criador": p.creator,
                "Data Criação": p.date,
                "Status": p.status === 'Completed' ? 'Concluído' : 'Em Andamento',
                "Total Ações": p.actions.length
            }));

            const ws = XLSX.utils.json_to_sheet(data);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "DiretorioProjetos");
            XLSX.writeFile(wb, `Brametal_Diretorio_Projetos.xlsx`);
        }

        // Painel de Indicadores e Gráficos (Fidelidade Cromática: Vermelho, Cinza e Branco)
        let chartsInstance = {};

        function updateDashboardDropdowns() {
            const idSelect = document.getElementById('filter-proj-id');
            const savedValue = idSelect.value;
            idSelect.innerHTML = '<option value="">Todos</option>';
            state.projects.forEach(p => {
                idSelect.innerHTML += `<option value="${p.id}">${p.id} - ${p.name.substring(0,20)}...</option>`;
            });
            idSelect.value = savedValue;
        }

        function updateDashboard() {
            const fStart = document.getElementById('filter-start-date').value;
            const fEnd = document.getElementById('filter-end-date').value;
            const fId = document.getElementById('filter-proj-id').value;
            const fCat = document.getElementById('filter-category').value;

            let filteredProjects = state.projects.filter(p => {
                if (fId && p.id !== fId) return false;
                if (fCat && p.category !== fCat) return false;
                if (fStart && p.date < fStart) return false;
                if (fEnd && p.date > fEnd) return false;
                return true;
            });

            let actions = [];
            filteredProjects.forEach(p => {
                p.actions.forEach(a => {
                    actions.push({ ...a, projectId: p.id, projectName: p.name, projectCategory: p.category });
                });
            });

            document.getElementById('dash-total-projects').innerText = filteredProjects.length;
            document.getElementById('dash-total-actions').innerText = actions.length;
            
            const delayedCount = actions.filter(a => a.kpi === 'Atrasado').length;
            const completedCount = actions.filter(a => a.status === 'Completed').length;
            
            document.getElementById('dash-delayed-actions').innerText = delayedCount;
            document.getElementById('dash-completion-rate').innerText = actions.length > 0 ? `${Math.round((completedCount / actions.length) * 100)}%` : '0%';

            let catMap = { "Melhorias de Sistemas": 0, "Melhorias de Processos": 0, "Estratégico": 0, "Operacional": 0, "Corretivas": 0, "Preventivas": 0 };
            filteredProjects.forEach(p => { if(catMap[p.category] !== undefined) catMap[p.category]++; });

            let statusMap = { 'Not Started': 0, 'In Progress': 0, 'Completed': 0 };
            actions.forEach(a => { if(statusMap[a.status] !== undefined) statusMap[a.status]++; });

            let kpiMap = { 'No Prazo': 0, 'Em Dia': 0, 'Atrasado': 0 };
            actions.forEach(a => { if(kpiMap[a.kpi] !== undefined) kpiMap[a.kpi]++; });

            let projActionsMap = {};
            actions.forEach(a => { projActionsMap[a.projectId] = (projActionsMap[a.projectId] || 0) + 1; });

            // Aplicação Estrita da Paleta Corporativa da Brametal nos Gráficos
            renderChart('chart-projects-category', 'bar', Object.keys(catMap), Object.values(catMap), [PALETTE.red, PALETTE.brandGray, PALETTE.mediumGray, PALETTE.lightGray, PALETTE.darkRed, '#4B5563'], 'Projetos');
            renderChart('chart-actions-status', 'doughnut', Object.keys(statusMap), Object.values(statusMap), [PALETTE.lightGray, PALETTE.brandGray, PALETTE.red], 'Ações');
            renderChart('chart-actions-kpi', 'pie', Object.keys(kpiMap), Object.values(kpiMap), [PALETTE.brandGray, PALETTE.ultraLightGray, PALETTE.red], 'Prazos');
            renderChart('chart-actions-project', 'bar', Object.keys(projActionsMap), Object.values(projActionsMap), [PALETTE.red], 'Ações');
        }

        function renderChart(canvasId, type, labels, data, colors, labelTitle) {
            if(chartsInstance[canvasId]) chartsInstance[canvasId].destroy();
            const ctx = document.getElementById(canvasId).getContext('2d');
            chartsInstance[canvasId] = new Chart(ctx, {
                type: type,
                data: {
                    labels: labels,
                    datasets: [{
                        label: labelTitle,
                        data: data,
                        backgroundColor: colors,
                        borderColor: '#E5E7EB',
                        borderWidth: 1
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { 
                        legend: { 
                            display: ['pie', 'doughnut'].includes(type), 
                            position: 'bottom',
                            labels: { font: { weight: 'bold' }, color: PALETTE.brandGray }
                        } 
                    }
                }
            });
        }
    </script>
</body>
</html>
