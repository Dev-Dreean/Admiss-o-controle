import React, { useState, useMemo, useEffect, useRef } from 'react';
import {
    UploadCloud, FileSpreadsheet, AlertCircle, Users, Search,
    X, ChevronRight, Briefcase, Download, Edit2, Save,
    LayoutDashboard, ListTodo, TableProperties, PlusCircle,
    BarChart2, PieChart as PieChartIcon, Trash2, CheckCircle2, XCircle, Send, PauseCircle,
    Undo2, Redo2, Check, ChevronLeft, Map, LogOut, Sliders
} from 'lucide-react';
import {
    PieChart, Pie, Cell, BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer
} from 'recharts';
import CryptoJS from 'crypto-js';
import * as XLSX from 'xlsx';

const COLORS = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899', '#14b8a6', '#f43f5e', '#84cc16'];
const STATUS_COLORS = {
    ABERTA: 'bg-red-100 text-red-800 border-red-300 font-bold',
    FECHADA: 'bg-green-100 text-green-800 border-green-300',
    ENCAMINHADA: 'bg-blue-100 text-blue-800 border-blue-300',
    CANCELADA: 'bg-gray-100 text-gray-800 border-gray-300',
    PAUSADA: 'bg-yellow-100 text-yellow-800 border-yellow-300',
};

const STATUS_ICON_MAP = {
    ABERTA: AlertCircle,
    FECHADA: CheckCircle2,
    ENCAMINHADA: Send,
    CANCELADA: XCircle,
    PAUSADA: PauseCircle,
};

const STATUS_ACCENT = {
    ABERTA: 'bg-red-500',
    FECHADA: 'bg-green-500',
    ENCAMINHADA: 'bg-blue-500',
    CANCELADA: 'bg-slate-400',
    PAUSADA: 'bg-yellow-400',
};

const GOOGLE_SHEETS_URL = 'https://docs.google.com/spreadsheets/d/1hmLkIX2B4rh6NDtJUXOhtjdXhddozqPs9uMTzaTeBsk/edit?usp=sharing';
const GOOGLE_SHEETS_CSV_EXPORT = 'https://docs.google.com/spreadsheets/d/1hmLkIX2B4rh6NDtJUXOhtjdXhddozqPs9uMTzaTeBsk/export?format=csv';
const AUTH_ACCESS_KEY = 'Plansul@2025';
const AUTH_SESSION_DURATION_MS = 24 * 60 * 60 * 1000;
const AUTH_DB_STORAGE_KEY = 'vagas_internal_account_excel_encrypted';
const AUTH_LEGACY_STORAGE_KEY = 'vagas_internal_account';
const AUTH_DB_SHEET_NAME = 'USUARIOS';
const AUTH_DB_SECRET = 'controle-admissoes-auth-db-v1';
const GRID_PAGE_SIZE = 200;
const SHEETS_PAGE_SIZE = 120;
const REQUIRED_FIELDS = ['Nome Subs', 'Status'];

const encodeBase64 = (arrayBuffer) => {
    const bytes = new Uint8Array(arrayBuffer);
    let binary = '';
    for (let i = 0; i < bytes.length; i += 1) binary += String.fromCharCode(bytes[i]);
    return window.btoa(binary);
};

const decodeBase64 = (base64) => {
    const binary = window.atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
    return bytes;
};

const createEncryptedExcelAuthDb = (account) => {
    if (!hasStoredAccount(account)) return '';

    const row = {
        Usuario: String(account.username || '').trim(),
        Senha: String(account.password || ''),
        CriadoEm: String(account.createdAt || ''),
        AtualizadoEm: String(account.updatedAt || ''),
        TutorialProgress: JSON.stringify(normalizeTutorialProgress(account.tutorialProgress || DEFAULT_TUTORIAL_PROGRESS)),
    };

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet([row]);
    XLSX.utils.book_append_sheet(workbook, worksheet, AUTH_DB_SHEET_NAME);
    const workbookBytes = XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });
    const workbookBase64 = encodeBase64(workbookBytes);
    return CryptoJS.AES.encrypt(workbookBase64, AUTH_DB_SECRET).toString();
};

const readEncryptedExcelAuthDb = (encryptedValue) => {
    if (!encryptedValue || typeof encryptedValue !== 'string') return null;

    try {
        const decrypted = CryptoJS.AES.decrypt(encryptedValue, AUTH_DB_SECRET).toString(CryptoJS.enc.Utf8);
        if (!decrypted) return null;

        const workbookBytes = decodeBase64(decrypted);
        const workbook = XLSX.read(workbookBytes, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        if (!sheetName) return null;

        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '', raw: false });
        if (!Array.isArray(rows) || rows.length === 0) return null;

        const latestRow = rows[0] || {};
        const account = {
            username: String(latestRow.Usuario || '').trim(),
            password: String(latestRow.Senha || ''),
            createdAt: String(latestRow.CriadoEm || ''),
            updatedAt: String(latestRow.AtualizadoEm || ''),
            tutorialProgress: normalizeTutorialProgress((() => {
                try {
                    return latestRow.TutorialProgress ? JSON.parse(String(latestRow.TutorialProgress)) : DEFAULT_TUTORIAL_PROGRESS;
                } catch (error) {
                    return DEFAULT_TUTORIAL_PROGRESS;
                }
            })()),
        };

        return hasStoredAccount(account) ? account : null;
    } catch (error) {
        return null;
    }
};

const safeGet = (obj, key) => String(obj[key] || '').trim().toUpperCase();

const parseDate = (dateStr) => {
    if (!dateStr || dateStr === '-') return null;
    const str = String(dateStr).trim().split(' ')[0];
    if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
        const [y, m, d] = str.split('-');
        return new Date(parseInt(y, 10), parseInt(m, 10) - 1, parseInt(d, 10));
    }
    if (/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(str)) {
        const [d, m, y] = str.split('/');
        return new Date(parseInt(y, 10), parseInt(m, 10) - 1, parseInt(d, 10));
    }
    const parsed = new Date(dateStr);
    if (!Number.isNaN(parsed.getTime())) return new Date(parsed.getTime() + parsed.getTimezoneOffset() * 60000);
    return null;
};

const getDaysDiff = (dateStr) => {
    const targetDate = parseDate(dateStr);
    if (!targetDate) return null;
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    targetDate.setHours(0, 0, 0, 0);
    return Math.ceil((targetDate - today) / (1000 * 60 * 60 * 24));
};

const getRowThermalClass = (dias, status, candidato, isInvalid) => {
    if (isInvalid) return 'bg-red-50 hover:bg-red-100 transition-colors border-l-4 border-red-600';
    if (['FECHADA', 'ENCAMINHADA', 'CANCELADA'].includes(status)) return 'hover:bg-slate-100 transition-colors bg-white opacity-70';
    if (candidato && candidato !== 'SEM COBERTURA') return 'hover:bg-blue-50 transition-colors bg-white';

    if (dias === null) return 'bg-slate-50 hover:bg-slate-100 transition-colors border-l-4 border-slate-200';
    if (dias <= 0) return 'bg-red-100 hover:bg-red-200 transition-colors border-l-4 border-red-500';
    if (dias <= 5) return 'bg-orange-100 hover:bg-orange-200 transition-colors border-l-4 border-orange-500';
    if (dias <= 15) return 'bg-yellow-50 hover:bg-yellow-100 transition-colors border-l-4 border-yellow-400';
    if (dias <= 30) return 'bg-lime-50 hover:bg-lime-100 transition-colors border-l-4 border-lime-400';
    return 'bg-green-50 hover:bg-green-100 transition-colors border-l-4 border-green-500';
};

const parseCSV = (text) => {
    const lines = text.split(/\r?\n/);
    if (lines.length === 0) return [];
    const splitLine = (line) => {
        const rowValues = [];
        let insideQuotes = false;
        let currentValue = '';
        for (let i = 0; i < line.length; i += 1) {
            const char = line[i];
            if (char === '"') insideQuotes = !insideQuotes;
            else if (char === ',' && !insideQuotes) {
                rowValues.push(currentValue.trim());
                currentValue = '';
            } else currentValue += char;
        }
        rowValues.push(currentValue.trim());
        return rowValues.map((v) => v.replace(/^"|"$/g, '').trim());
    };
    const headers = splitLine(lines[0]);
    if (headers.length === 0 || !headers[0]) return [];
    const result = [];
    for (let i = 1; i < lines.length; i += 1) {
        if (!lines[i].trim()) continue;
        const rowValues = splitLine(lines[i]);
        const obj = {};
        let hasData = false;
        headers.forEach((header, index) => {
            if (header) {
                obj[header] = rowValues[index] || '';
                if (obj[header]) hasData = true;
            }
        });
        if (hasData) result.push(obj);
    }
    return result;
};

function useLocalStorage(key, initialValue) {
    const [storedValue, setStoredValue] = useState(() => {
        if (typeof window === 'undefined') return initialValue;
        try {
            const item = window.localStorage.getItem(key);
            return item ? JSON.parse(item) : initialValue;
        } catch (error) {
            return initialValue;
        }
    });

    const setValue = (value) => {
        try {
            const valueToStore = value instanceof Function ? value(storedValue) : value;
            setStoredValue(valueToStore);
            if (typeof window !== 'undefined') {
                window.localStorage.setItem(key, JSON.stringify(valueToStore));
            }
        } catch (error) {
            // no-op
        }
    };

    return [storedValue, setValue];
}

const getFirstName = (fullName) => {
    const normalized = String(fullName || '').trim();
    if (!normalized) return 'USUARIO';
    return normalized.split(/\s+/)[0].toUpperCase();
};

const normalizeCredentialText = (value) => String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();

const matchesSavedUsername = (inputValue, savedUsername) => {
    const normalizedInput = normalizeCredentialText(inputValue);
    const normalizedSaved = normalizeCredentialText(savedUsername);

    if (!normalizedInput || !normalizedSaved) return false;
    if (normalizedInput === normalizedSaved) return true;

    return normalizedInput === normalizeCredentialText(getFirstName(savedUsername));
};

const getAuthSessionExpiryTimestamp = (session) => {
    const startedAt = Date.parse(String(session?.at || ''));
    if (Number.isNaN(startedAt)) return null;
    return startedAt + AUTH_SESSION_DURATION_MS;
};

const hasActiveAuthSession = (session, account) => {
    if (!account || !session?.authenticated || !matchesSavedUsername(session.username, account.username)) return false;
    const expiryTimestamp = getAuthSessionExpiryTimestamp(session);
    return Boolean(expiryTimestamp && expiryTimestamp > Date.now());
};

const hasStoredAccount = (account) => Boolean(
    account
    && typeof account === 'object'
    && String(account.username || '').trim()
    && String(account.password || '').trim(),
);

const PRAZO_KEYS = [
    'Prazo',
    'Inicio Situacao',
    'Inicio Situação',
    'Início Situação',
    'Data Prazo',
    'Data de Prazo',
    'Data Inicio',
];

const getRowValue = (row, possibleKeys) => {
    if (!row || typeof row !== 'object') return '';

    for (const key of possibleKeys) {
        const directValue = row[key];
        if (directValue !== undefined && String(directValue).trim() !== '') return directValue;
    }

    const normalizedKeys = new Set(possibleKeys.map((key) => normalizeCredentialText(key)));
    for (const [rowKey, rowValue] of Object.entries(row)) {
        if (normalizedKeys.has(normalizeCredentialText(rowKey)) && String(rowValue).trim() !== '') {
            return rowValue;
        }
    }

    return '';
};

const getRowKey = (row, possibleKeys) => {
    if (!row || typeof row !== 'object') return possibleKeys[0] || '';

    for (const key of possibleKeys) {
        if (Object.prototype.hasOwnProperty.call(row, key)) return key;
    }

    const normalizedKeys = new Set(possibleKeys.map((key) => normalizeCredentialText(key)));
    for (const rowKey of Object.keys(row)) {
        if (normalizedKeys.has(normalizeCredentialText(rowKey))) return rowKey;
    }

    return possibleKeys[0] || '';
};

const getPrazoValue = (row) => getRowValue(row, PRAZO_KEYS);
const getPrazoKey = (row) => getRowKey(row, PRAZO_KEYS);
const SHEETS_PRIORITY_COLUMNS = ['Status', 'Nome Subs', 'CARGO', 'Candidato', 'Contato Candidato', 'NRE / MUNICIPIO', 'Motivo', 'OBS:'];

const DEFAULT_TUTORIAL_PROGRESS = Object.freeze({
    TABELA: false,
    DASHBOARD: false,
    SHEETS: false,
    NEW_FEATURES_TABLE: false,
    NEW_FEATURES_SHEETS: false,
    NEW_FEATURES_SHEETS_V2: false,
    PATCH_NOTES_VIEWED: false,
});

const normalizeTutorialProgress = (value) => ({
    TABELA: Boolean(value?.TABELA),
    DASHBOARD: Boolean(value?.DASHBOARD),
    SHEETS: Boolean(value?.SHEETS),
    NEW_FEATURES_TABLE: Boolean(value?.NEW_FEATURES_TABLE),
    NEW_FEATURES_SHEETS: Boolean(value?.NEW_FEATURES_SHEETS),
    NEW_FEATURES_SHEETS_V2: Boolean(value?.NEW_FEATURES_SHEETS_V2),
    PATCH_NOTES_VIEWED: Boolean(value?.PATCH_NOTES_VIEWED),
});

const TUTORIAL_STEPS = {
    TABELA: [
        {
            id: 'table-navigation',
            target: 'tour-tabs',
            title: 'Navegacao',
            desc: 'Aqui voce troca entre a tabela, os graficos e a planilha, sempre seguindo o fluxo principal do sistema.',
            icon: <LayoutDashboard className="w-8 h-8 text-indigo-500" />,
        },
        {
            id: 'table-actions',
            target: 'tour-header-actions',
            title: 'Acoes rapidas',
            desc: 'Neste bloco ficam desfazer, refazer, importar, exportar e atualizar. Use essas acoes para corrigir, trazer novos dados ou baixar a base atual.',
            icon: <UploadCloud className="w-8 h-8 text-indigo-500" />,
        },
        {
            id: 'table-record-count',
            target: 'tour-record-count',
            title: 'Exibidos',
            desc: 'Este contador mostra quantas linhas estao aparecendo agora, ja considerando busca e filtros ativos.',
            icon: <FileSpreadsheet className="w-8 h-8 text-indigo-500" />,
        },
        {
            id: 'table-filters',
            target: 'tour-filters-controls',
            title: 'Filtros',
            desc: 'Aqui voce usa busca, status, municipio, prazo e o botao Corrigir Erros para chegar mais rapido no grupo de vagas que precisa analisar.',
            icon: <Search className="w-8 h-8 text-orange-500" />,
        },
        {
            id: 'table-head',
            target: 'tour-table-head',
            title: 'Cabecalho da tabela',
            desc: 'O cabecalho resume as informacoes da grade: status rapido, vaga, candidato, prazo e ficha.',
            icon: <ListTodo className="w-8 h-8 text-indigo-500" />,
        },
        {
            id: 'table-grid',
            target: 'tour-table',
            title: 'Grid de acompanhamento',
            desc: 'Nesta grade cada linha representa uma vaga. Aqui voce acompanha a situacao geral sem precisar abrir tudo de uma vez.',
            icon: <TableProperties className="w-8 h-8 text-sky-500" />,
        },
        {
            id: 'table-record-button',
            target: 'tour-record-open-btn',
            title: 'Fichas',
            desc: 'Use este botao da ultima coluna para abrir a ficha completa da vaga selecionada.',
            icon: <Edit2 className="w-8 h-8 text-blue-500" />,
        },
        {
            id: 'table-record-fields',
            target: 'tour-record-edit-fields',
            title: 'Alteracoes',
            desc: 'Dentro da ficha, esta area concentra os campos editaveis para atualizar status, profissional, candidato, contato, municipio, prazo e observacoes.',
            icon: <Check className="w-8 h-8 text-emerald-500" />,
        },
        {
            id: 'table-record-save',
            target: 'tour-record-save-btn',
            title: 'Salvar alteracoes',
            desc: 'Depois de revisar a ficha, salve para aplicar as mudancas na linha e no painel.',
            icon: <Save className="w-8 h-8 text-green-600" />,
        },
    ],
    DASHBOARD: [
        {
            id: 'dashboard-tab',
            target: 'tour-tab-dashboard',
            title: 'Tela de graficos',
            desc: 'Ao clicar aqui, voce abre a area com os indicadores visuais da operacao.',
            icon: <BarChart2 className="w-8 h-8 text-indigo-500" />,
        },
        {
            id: 'dashboard-panel',
            target: 'tour-dashboard-panel',
            title: 'Resumo visual',
            desc: 'Nesta tela voce bate o olho nos totais, nos status e nos principais motivos em grafico.',
            icon: <LayoutDashboard className="w-8 h-8 text-sky-500" />,
        },
        {
            id: 'dashboard-create-chart',
            target: 'tour-create-chart-btn',
            title: 'Criar grafico',
            desc: 'Este botao abre o criador de grafico para montar novas visoes do painel.',
            icon: <PlusCircle className="w-8 h-8 text-blue-500" />,
        },
        {
            id: 'dashboard-chart-modal',
            target: 'tour-chart-modal',
            title: 'Modal de grafico',
            desc: 'Aqui voce define o titulo, a coluna usada e o tipo do grafico antes de salvar no dashboard.',
            icon: <PieChartIcon className="w-8 h-8 text-indigo-500" />,
        },
    ],
    SHEETS: [
        {
            id: 'sheets-tab',
            target: 'tour-tab-sheets',
            title: 'Modo planilha',
            desc: 'Esta aba abre uma visao parecida com planilha para editar linha por linha de forma rapida.',
            icon: <TableProperties className="w-8 h-8 text-green-500" />,
        },
        {
            id: 'sheets-head',
            target: 'tour-sheets-head',
            title: 'Cabecalho da planilha',
            desc: 'O cabecalho agrupa status, profissional, candidato cobertura, telefone, municipio e observacoes para orientar a edicao da linha.',
            icon: <ListTodo className="w-8 h-8 text-indigo-500" />,
        },
        {
            id: 'sheets-table',
            target: 'tour-sheets-table',
            title: 'Linhas editaveis',
            desc: 'Todas as alteracoes feitas nas linhas desta planilha entram direto no painel e ajudam a manter a base atualizada.',
            icon: <Check className="w-8 h-8 text-emerald-500" />,
        },
    ],
    NEW_FEATURES_TABLE: [
        {
            id: 'new-features-quick-add',
            target: 'tour-quick-add-btn',
            title: 'Novo cadastro rapido',
            desc: 'Use este botao para inserir uma pessoa nova no sistema em segundos, sem sair da tela atual.',
            icon: <PlusCircle className="w-8 h-8 text-emerald-500" />,
        },
        {
            id: 'new-features-quick-add-modal',
            target: 'tour-quick-add-modal',
            title: 'Modal de cadastro',
            desc: 'Neste modal voce cadastra todos os campos obrigatorios e complementares em um formulario unico.',
            icon: <Edit2 className="w-8 h-8 text-indigo-500" />,
        },
    ],
    NEW_FEATURES_SHEETS_V2: [
        {
            id: 'new-features-dynamic-filters',
            target: 'tour-sheets-dynamic-filters',
            title: 'Filtros dinamicos',
            desc: 'Aqui voce filtra a planilha com busca global e um filtro rapido por coluna, sem poluicao visual.',
            icon: <Sliders className="w-8 h-8 text-green-600" />,
        },
        {
            id: 'new-features-columns-panel',
            target: 'tour-sheets-columns-panel-btn',
            title: 'Colunas inteligentes',
            desc: 'No botao Colunas voce abre o painel para escolher o que aparece na planilha e ajustar o layout para o seu fluxo.',
            icon: <TableProperties className="w-8 h-8 text-emerald-600" />,
        },
        {
            id: 'new-features-columns-toggle',
            target: 'tour-sheets-columns-first-toggle',
            title: 'Desmarcar e remarcar coluna',
            desc: 'Clique no checkbox de uma coluna para desmarcar: ela some da planilha. Clique de novo para remarcar: ela volta. Assim voce personaliza do jeito que preferir.',
            icon: <Check className="w-8 h-8 text-indigo-500" />,
        },
        {
            id: 'new-features-actions-delete',
            target: 'tour-sheets-actions-col',
            title: 'Acoes e exclusao',
            desc: 'A coluna Acoes fica no lado direito da planilha. Role para a direita para chegar nela e excluir uma linha pela lixeira.',
            icon: <Trash2 className="w-8 h-8 text-rose-600" />,
        },
        {
            id: 'new-features-horizontal-scroll',
            target: 'tour-sheets-horizontal-scroll',
            title: 'Rolagem lateral rapida',
            desc: 'Para navegar nas laterais com mais controle, segure Shift e use o scroll do mouse: isso move horizontalmente a planilha.',
            icon: <ChevronRight className="w-8 h-8 text-indigo-500" />,
        },
    ],
};

const StatusBadge = ({ status, size = 'sm' }) => {
    const normalized = String(status || '').toUpperCase().trim();
    const colorClass = STATUS_COLORS[normalized] || 'bg-slate-100 text-slate-600 border-slate-200';
    const IconComponent = STATUS_ICON_MAP[normalized];
    return (
        <span className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full border ${colorClass} ${size === 'sm' ? 'text-xs' : 'text-sm'} font-bold`}>
            {IconComponent && <IconComponent className={size === 'sm' ? 'w-3 h-3' : 'w-4 h-4'} />}
            {normalized || '—'}
        </span>
    );
};

const PatchNotesModal = ({ onClose, onStartTableTutorial, onStartSheetsTutorial }) => {
    return (
        <div className="fixed inset-0 z-[200] bg-slate-900/60 flex items-center justify-center p-4 backdrop-blur-sm animate-in fade-in duration-200">
            <div className="bg-white rounded-3xl shadow-2xl w-full max-w-2xl max-h-[90vh] flex flex-col overflow-hidden animate-in zoom-in-95 duration-300">
                <div className="px-8 py-6 border-b bg-gradient-to-r from-indigo-50 to-purple-50 flex items-center justify-between">
                    <div className="flex items-center gap-3">
                        <div className="bg-indigo-600 p-3 rounded-2xl">
                            <BarChart2 className="w-6 h-6 text-white" />
                        </div>
                        <div>
                            <h2 className="text-2xl font-black text-slate-800">Sistema Atualizado</h2>
                            <p className="text-sm text-slate-500 mt-0.5">Veja as novidades implementadas</p>
                        </div>
                    </div>
                    <button onClick={onClose} className="p-2 text-slate-400 hover:bg-slate-200 rounded-xl transition-colors" type="button">
                        <X className="w-5 h-5" />
                    </button>
                </div>

                <div className="p-8 overflow-y-auto flex-1 space-y-6">
                    <div className="space-y-4">
                        <div className="flex gap-4">
                            <div className="flex-shrink-0">
                                <div className="flex items-center justify-center h-10 w-10 rounded-full bg-emerald-100">
                                    <PlusCircle className="h-6 w-6 text-emerald-600" />
                                </div>
                            </div>
                            <div className="flex-1">
                                <h3 className="text-lg font-bold text-slate-800">Cadastro Rápido</h3>
                                <p className="text-slate-600 text-sm mt-1">Novo botão para inserir pessoas no sistema em segundos, sem sair da tela atual. Formulário completo com todos os campos em uma única aba.</p>
                            </div>
                        </div>

                        <div className="flex gap-4">
                            <div className="flex-shrink-0">
                                <div className="flex items-center justify-center h-10 w-10 rounded-full bg-blue-100">
                                    <Edit2 className="h-6 w-6 text-blue-600" />
                                </div>
                            </div>
                            <div className="flex-1">
                                <h3 className="text-lg font-bold text-slate-800">Edição completa de Fichas</h3>
                                <p className="text-slate-600 text-sm mt-1">Agora você pode editar <strong>todos os campos</strong> da ficha diretamente no modal. Modo anterior era limitado a apenas alguns campos.</p>
                            </div>
                        </div>

                        <div className="flex gap-4">
                            <div className="flex-shrink-0">
                                <div className="flex items-center justify-center h-10 w-10 rounded-full bg-orange-100">
                                    <Sliders className="h-6 w-6 text-orange-600" />
                                </div>
                            </div>
                            <div className="flex-1">
                                <h3 className="text-lg font-bold text-slate-800">Filtros Simplificados</h3>
                                <p className="text-slate-600 text-sm mt-1">Interface minimalista para pesquisar. Busca global + filtro rápido por coluna. Sem excesso de informações, ainda mais produtivo.</p>
                            </div>
                        </div>

                        <div className="flex gap-4">
                            <div className="flex-shrink-0">
                                <div className="flex items-center justify-center h-10 w-10 rounded-full bg-purple-100">
                                    <Briefcase className="h-6 w-6 text-purple-600" />
                                </div>
                            </div>
                            <div className="flex-1">
                                <h3 className="text-lg font-bold text-slate-800">Atalhos de Teclado</h3>
                                <p className="text-slate-600 text-sm mt-1"><kbd className="px-2 py-1 bg-slate-100 border border-slate-300 rounded text-xs font-mono">C</kbd> abre cadastro rápido • <kbd className="px-2 py-1 bg-slate-100 border border-slate-300 rounded text-xs font-mono">F</kbd> foca busca global</p>
                            </div>
                        </div>

                        <div className="flex gap-4">
                            <div className="flex-shrink-0">
                                <div className="flex items-center justify-center h-10 w-10 rounded-full bg-red-100">
                                    <Trash2 className="h-6 w-6 text-red-600" />
                                </div>
                            </div>
                            <div className="flex-1">
                                <h3 className="text-lg font-bold text-slate-800">Remocao de Registro</h3>
                                <p className="text-slate-600 text-sm mt-1">Agora e possivel excluir direto pela coluna de acoes na tabela e tambem dentro da ficha completa.</p>
                            </div>
                        </div>
                    </div>

                    <div className="bg-indigo-50 border border-indigo-200 rounded-2xl p-4">
                        <p className="text-sm text-indigo-900">
                            <span className="font-bold">Dica:</span> Escolha o atalho abaixo para abrir o tutorial pratico da tela desejada.
                        </p>
                    </div>
                </div>

                <div className="px-8 py-4 border-t bg-slate-50 flex flex-wrap justify-end gap-3">
                    <button onClick={onClose} className="px-6 py-2.5 text-slate-700 font-bold hover:bg-slate-200 rounded-xl transition-colors" type="button">
                        Agora nao
                    </button>
                    <button onClick={onStartSheetsTutorial} className="px-6 py-2.5 bg-emerald-600 text-white font-bold rounded-xl hover:bg-emerald-700 shadow-md active:scale-95 transition-all inline-flex items-center gap-2" type="button">
                        <ChevronRight className="w-4 h-4" />
                        Tutorial Planilha
                    </button>
                    <button onClick={onStartTableTutorial} className="px-6 py-2.5 bg-indigo-600 text-white font-bold rounded-xl hover:bg-indigo-700 shadow-md active:scale-95 transition-all inline-flex items-center gap-2" type="button">
                        <ChevronRight className="w-4 h-4" />
                        Tutorial Tabela
                    </button>
                </div>
            </div>
        </div>
    );
};

const TUTORIAL_RECORD_MODAL_STEP_IDS = new Set(['table-record-fields', 'table-record-save']);
const TUTORIAL_RECORD_EDIT_STEP_IDS = new Set(['table-record-fields', 'table-record-save']);
const TUTORIAL_CHART_MODAL_STEP_IDS = new Set(['dashboard-chart-modal']);

const WelcomeOverlay = ({ onFinish, userName }) => {
    const [stage, setStage] = useState('blank');
    const firstName = getFirstName(userName);

    useEffect(() => {
        const t1 = setTimeout(() => setStage('purple'), 1000);
        const t2 = setTimeout(() => setStage('welcome'), 2000);
        const t3 = setTimeout(() => setStage('system'), 3200);
        const t4 = setTimeout(() => setStage('out'), 4300);
        const t5 = setTimeout(() => onFinish(), 4900);
        return () => {
            clearTimeout(t1);
            clearTimeout(t2);
            clearTimeout(t3);
            clearTimeout(t4);
            clearTimeout(t5);
        };
    }, [onFinish]);

    return (
        <div className={`fixed inset-0 z-[300] pointer-events-none transition-opacity duration-700 ${stage === 'out' ? 'opacity-0' : 'opacity-100'}`}>
            <div className="absolute inset-0 bg-slate-950" />

            <div
                className={`absolute inset-0 transition-opacity duration-700 ${['purple', 'welcome', 'system'].includes(stage) ? 'opacity-100' : 'opacity-0'}`}
                style={{
                    background: 'radial-gradient(circle at 20% 15%, rgba(216,180,254,0.68), transparent 40%), radial-gradient(circle at 80% 90%, rgba(196,181,253,0.45), transparent 45%), linear-gradient(135deg, rgba(237,233,254,0.95), rgba(224,231,255,0.9))',
                    filter: 'blur(10px)',
                }}
            />

            <div className="absolute inset-0 flex items-center justify-center px-6">
                <h1
                    className={`text-center text-5xl md:text-7xl font-black tracking-tight text-violet-950 transition-all duration-500 ${stage === 'welcome' ? 'opacity-100 translate-y-0' : 'opacity-0 -translate-y-2'}`}
                >
                    {`BOAS-VINDAS, ${firstName}`}
                </h1>

                <h2
                    className={`absolute text-center text-4xl md:text-6xl font-black tracking-tight text-violet-950 transition-all duration-500 ${stage === 'system' ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-2'}`}
                >
                    RECRUTAMENTO
                </h2>
            </div>
        </div>
    );
};

const WalkthroughTour = ({ section, steps, onComplete, onStepChange }) => {
    const [step, setStep] = useState(0);
    const [rect, setRect] = useState(null);
    const [dontShowAgain, setDontShowAgain] = useState(false);
    const [dialogSize, setDialogSize] = useState({ width: 360, height: 320 });
    const dialogRef = useRef(null);
    const currentStep = steps[step];
    const progressPercent = steps.length > 1 ? ((step + 1) / steps.length) * 100 : 100;

    useEffect(() => {
        setStep(0);
        setRect(null);
        setDontShowAgain(false);
    }, [section]);

    useEffect(() => {
        if (onStepChange) onStepChange(currentStep || null);
    }, [currentStep, onStepChange]);

    useEffect(() => {
        const measureDialog = () => {
            if (!dialogRef.current) return;
            const { width, height } = dialogRef.current.getBoundingClientRect();
            setDialogSize((previous) => (
                Math.abs(previous.width - width) > 1 || Math.abs(previous.height - height) > 1
                    ? { width, height }
                    : previous
            ));
        };

        const timer = setTimeout(measureDialog, 0);
        window.addEventListener('resize', measureDialog);
        return () => {
            clearTimeout(timer);
            window.removeEventListener('resize', measureDialog);
        };
    }, [section, step]);

    useEffect(() => {
        if (!currentStep) return undefined;

        const updateRect = () => {
            const el = document.getElementById(currentStep.target);
            if (el) {
                const r = el.getBoundingClientRect();
                if (r.top < 80 || r.bottom > window.innerHeight - 80) {
                    el.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
                    setTimeout(() => {
                        const newEl = document.getElementById(currentStep.target);
                        if (newEl) setRect(newEl.getBoundingClientRect());
                    }, 420);
                } else {
                    setRect(r);
                }
            } else {
                setRect(null);
            }
        };

        const timer = setTimeout(updateRect, 450);
        window.addEventListener('resize', updateRect);
        return () => {
            clearTimeout(timer);
            window.removeEventListener('resize', updateRect);
        };
    }, [currentStep]);

    const dialogStyle = useMemo(() => {
        const viewportPadding = 16;
        const viewportWidth = typeof window !== 'undefined' ? window.innerWidth : 1280;
        const viewportHeight = typeof window !== 'undefined' ? window.innerHeight : 720;
        const fallbackWidth = Math.min(420, viewportWidth - (viewportPadding * 2));
        const dialogWidth = Math.min(dialogSize.width || fallbackWidth, fallbackWidth);
        const dialogHeight = Math.min(dialogSize.height || 320, viewportHeight - (viewportPadding * 2));

        if (!rect) {
            return {
                width: dialogWidth,
                maxHeight: viewportHeight - (viewportPadding * 2),
                top: '50%',
                left: '50%',
                transform: 'translate(-50%, -50%)',
            };
        }

        const preferredLeft = rect.left + (rect.width / 2) - (dialogWidth / 2);
        const canShowBelow = rect.bottom + 20 + dialogHeight <= viewportHeight - viewportPadding;
        const preferredTop = canShowBelow ? rect.bottom + 20 : rect.top - dialogHeight - 20;
        const maxLeft = viewportWidth - dialogWidth - viewportPadding;
        const maxTop = viewportHeight - dialogHeight - viewportPadding;

        return {
            width: dialogWidth,
            maxHeight: viewportHeight - (viewportPadding * 2),
            top: Math.max(viewportPadding, Math.min(preferredTop, maxTop)),
            left: Math.max(viewportPadding, Math.min(preferredLeft, maxLeft)),
        };
    }, [dialogSize.height, dialogSize.width, rect]);

    if (!currentStep) return null;

    return (
        <div className="fixed inset-0 z-[150] pointer-events-auto transition-opacity duration-500 animate-in fade-in">
            {rect ? (
                <div
                    className="absolute transition-all duration-500 pointer-events-none"
                    style={{
                        top: rect.top - 12,
                        left: rect.left - 12,
                        width: rect.width + 24,
                        height: rect.height + 24,
                        borderRadius: '16px',
                        boxShadow: '0 0 0 9999px rgba(15,23,42,0.85)',
                    }}
                />
            ) : (
                <div className="absolute inset-0 bg-slate-900/85 pointer-events-none" />
            )}

            <div
                ref={dialogRef}
                className="absolute bg-white rounded-3xl p-5 sm:p-6 shadow-[0_0_40px_rgba(0,0,0,0.3)] transition-all duration-500 animate-in zoom-in-95 overflow-y-auto overflow-x-hidden"
                style={dialogStyle}
            >
                <button onClick={() => onComplete(false)} className="absolute top-4 right-4 text-slate-400 hover:text-slate-800 transition-colors" type="button"><X className="w-5 h-5" /></button>

                <div className="flex items-center gap-3 mb-4">
                    <div className="bg-indigo-50 p-2 rounded-2xl">{currentStep.icon}</div>
                    <div>
                        <p className="text-[11px] font-bold uppercase tracking-[0.22em] text-slate-400">Passo {step + 1} de {steps.length}</p>
                        <h3 className="text-xl font-bold text-slate-800 pr-8">{currentStep.title}</h3>
                    </div>
                </div>

                <p className="text-slate-600 text-sm mb-8 leading-relaxed">{currentStep.desc}</p>

                <div className="mt-auto space-y-4">
                    <div className="space-y-2">
                        <div className="flex items-center justify-between gap-3 text-[11px] font-bold uppercase tracking-[0.22em] text-slate-400">
                            <span>Andamento</span>
                            <span>{step + 1}/{steps.length}</span>
                        </div>
                        <div className="h-2 overflow-hidden rounded-full bg-slate-200">
                            <div className="h-full rounded-full bg-indigo-600 transition-all duration-500" style={{ width: `${progressPercent}%` }} />
                        </div>
                    </div>
                    <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                        <div className="min-h-[24px]">
                            {step === steps.length - 1 && (
                                <label className="flex items-center gap-2 cursor-pointer group">
                                    <input type="checkbox" checked={dontShowAgain} onChange={(e) => setDontShowAgain(e.target.checked)} className="rounded text-indigo-600 focus:ring-indigo-500 w-4 h-4 cursor-pointer" />
                                    <span className="text-[11px] font-medium text-slate-400 group-hover:text-slate-600 transition-colors">Nao exibir este tutorial novamente</span>
                                </label>
                            )}
                        </div>
                        <div className="flex items-center justify-end gap-2">
                            {step > 0 && (
                                <button
                                    onClick={() => setStep((s) => Math.max(0, s - 1))}
                                    className="inline-flex items-center justify-center gap-2 bg-slate-100 hover:bg-slate-200 text-slate-700 min-h-[44px] px-4 rounded-full shadow-sm active:scale-95 transition-all shrink-0"
                                    type="button"
                                    aria-label="Voltar um passo"
                                >
                                    <ChevronLeft className="w-4 h-4" />
                                    <span className="text-sm font-semibold">Voltar</span>
                                </button>
                            )}
                            <button
                                onClick={() => {
                                    if (step < steps.length - 1) setStep((s) => s + 1);
                                    else onComplete(dontShowAgain);
                                }}
                                className="inline-flex items-center justify-center gap-2 bg-indigo-600 hover:bg-indigo-700 text-white min-h-[44px] px-5 rounded-full text-sm font-bold shadow-md active:scale-95 transition-all shrink-0"
                                type="button"
                            >
                                <span>{step < steps.length - 1 ? 'Avancar' : 'Concluir'}</span>
                                <ChevronRight className="w-4 h-4" />
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    );
};

const LoginOverlay = ({ account, onCreateAccount, onLogin, notice }) => {
    const hasAccount = hasStoredAccount(account);
    const [mode, setMode] = useState(hasAccount ? 'login' : 'create');
    const [createForm, setCreateForm] = useState({ username: '', password: '', confirmPassword: '', accessKey: '' });
    const [loginUser, setLoginUser] = useState('');
    const [loginPass, setLoginPass] = useState('');
    const [error, setError] = useState('');

    useEffect(() => {
        setMode(hasAccount ? 'login' : 'create');
        setLoginUser(hasAccount ? account.username : '');
        setLoginPass('');
        setCreateForm({ username: '', password: '', confirmPassword: '', accessKey: '' });
        setError('');
    }, [account, hasAccount]);

    const switchMode = (nextMode) => {
        if (nextMode === 'login' && !hasAccount) {
            setError('Ainda nao existe uma conta criada neste navegador.');
            return;
        }

        setMode(nextMode);
        setError('');
        if (nextMode === 'login' && hasAccount) {
            setLoginUser(account.username);
            setLoginPass('');
        }
    };

    const handleCreate = (e) => {
        e.preventDefault();
        if (!createForm.username.trim() || !createForm.password.trim()) {
            setError('Preencha usuario e senha para criar o acesso.');
            return;
        }
        if (createForm.password !== createForm.confirmPassword) {
            setError('A confirmacao da senha precisa ser igual a senha.');
            return;
        }
        if (createForm.accessKey !== AUTH_ACCESS_KEY) {
            setError('Palavra-chave invalida. O acesso nao foi liberado.');
            return;
        }

        const now = new Date().toISOString();
        onCreateAccount({
            username: createForm.username.trim(),
            password: createForm.password,
            createdAt: account?.createdAt || now,
            updatedAt: now,
        });
    };

    const handleLogin = (e) => {
        e.preventDefault();
        if (!loginUser.trim() || !loginPass.trim()) {
            setError('Digite usuario e senha para entrar.');
            return;
        }
        const isUserValid = matchesSavedUsername(loginUser, account?.username);
        const savedPassword = String(account?.password || '');
        const currentPassword = String(loginPass || '');
        const isPasswordValid = savedPassword && (
            currentPassword === savedPassword
            || currentPassword.trim() === savedPassword.trim()
        );
        if (!isUserValid || !isPasswordValid) {
            setError('Credenciais invalidas. Use o usuario cadastrado ou apenas o primeiro nome.');
            return;
        }
        setError('');
        onLogin();
    };

    return (
        <div className="fixed inset-0 z-[260] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm">
            <div className="w-full max-w-md bg-white rounded-3xl border border-slate-200 shadow-2xl p-7">
                <div className="flex items-center gap-3 mb-5">
                    <div className="bg-indigo-600 p-2 rounded-2xl shadow-sm"><Briefcase className="w-5 h-5 text-white" /></div>
                    <div>
                        <h2 className="text-2xl font-black text-slate-800">Acesso ao sistema</h2>
                        <p className="text-sm text-slate-500">Crie seu acesso com a palavra-chave ou entre com o login ja salvo.</p>
                    </div>
                </div>

                <div className="grid grid-cols-2 gap-2 p-1 mb-6 rounded-2xl bg-slate-100">
                    <button
                        onClick={() => switchMode('create')}
                        className={`px-4 py-2.5 rounded-xl text-sm font-bold transition-all ${mode === 'create' ? 'bg-white text-indigo-700 shadow-sm' : 'text-slate-500 hover:text-slate-700'}`}
                        type="button"
                    >
                        Criar conta
                    </button>
                    <button
                        onClick={() => switchMode('login')}
                        className={`px-4 py-2.5 rounded-xl text-sm font-bold transition-all ${mode === 'login' ? 'bg-white text-indigo-700 shadow-sm' : 'text-slate-500 hover:text-slate-700'} ${!hasAccount ? 'opacity-50 cursor-not-allowed' : ''}`}
                        type="button"
                    >
                        Ja tenho login
                    </button>
                </div>

                {notice && <div className="mb-5 rounded-2xl border border-amber-200 bg-amber-50 px-4 py-3 text-sm font-medium text-amber-800">{notice}</div>}

                {mode === 'create' ? (
                    <div>
                        <p className="text-sm text-slate-500 mb-5">
                            {hasAccount
                                ? 'Se precisar trocar o acesso, crie um novo login com a palavra-chave. Os dados do painel continuam salvos.'
                                : 'Esta e a tela inicial para criar a primeira conta deste navegador.'}
                        </p>

                        <form className="space-y-4" onSubmit={handleCreate}>
                            <input
                                value={createForm.username}
                                onChange={(e) => {
                                    setCreateForm((current) => ({ ...current, username: e.target.value }));
                                    if (error) setError('');
                                }}
                                className="w-full px-4 py-2.5 rounded-xl border border-slate-300 focus:ring-2 focus:ring-indigo-500"
                                placeholder="Usuario"
                                autoComplete="username"
                            />
                            <input
                                type="password"
                                value={createForm.password}
                                onChange={(e) => {
                                    setCreateForm((current) => ({ ...current, password: e.target.value }));
                                    if (error) setError('');
                                }}
                                className="w-full px-4 py-2.5 rounded-xl border border-slate-300 focus:ring-2 focus:ring-indigo-500"
                                placeholder="Senha"
                                autoComplete="new-password"
                            />
                            <input
                                type="password"
                                value={createForm.confirmPassword}
                                onChange={(e) => {
                                    setCreateForm((current) => ({ ...current, confirmPassword: e.target.value }));
                                    if (error) setError('');
                                }}
                                className="w-full px-4 py-2.5 rounded-xl border border-slate-300 focus:ring-2 focus:ring-indigo-500"
                                placeholder="Confirmar senha"
                                autoComplete="new-password"
                            />
                            <input
                                type="password"
                                value={createForm.accessKey}
                                onChange={(e) => {
                                    setCreateForm((current) => ({ ...current, accessKey: e.target.value }));
                                    if (error) setError('');
                                }}
                                className="w-full px-4 py-2.5 rounded-xl border border-slate-300 focus:ring-2 focus:ring-indigo-500"
                                placeholder="Palavra-chave de autorizacao"
                            />
                            <button type="submit" className="w-full bg-indigo-600 hover:bg-indigo-700 text-white py-2.5 rounded-xl font-bold transition-colors">
                                {hasAccount ? 'Salvar novo acesso' : 'Criar conta e entrar'}
                            </button>
                        </form>

                        <p className="text-xs text-slate-400 mt-4 leading-relaxed">
                            Palavra-chave para criar ou trocar o acesso: <span className="font-bold text-slate-600">{AUTH_ACCESS_KEY}</span>.
                            Criar um novo acesso nao apaga os dados do painel.
                        </p>
                    </div>
                ) : (
                    <div>
                        <p className="text-sm text-slate-500 mb-5">Use o usuario e a senha cadastrados para liberar o painel.</p>
                        {hasAccount && <p className="text-xs text-slate-400 mb-4">Usuario salvo: <span className="font-bold text-slate-600">{account.username}</span>. Voce pode entrar com o nome completo ou apenas o primeiro nome.</p>}

                        <form className="space-y-4" onSubmit={handleLogin}>
                            <input
                                value={loginUser}
                                onChange={(e) => {
                                    setLoginUser(e.target.value);
                                    if (error) setError('');
                                }}
                                className="w-full px-4 py-2.5 rounded-xl border border-slate-300 focus:ring-2 focus:ring-indigo-500"
                                placeholder="Usuario"
                                autoComplete="username"
                            />
                            <input
                                type="password"
                                value={loginPass}
                                onChange={(e) => {
                                    setLoginPass(e.target.value);
                                    if (error) setError('');
                                }}
                                className="w-full px-4 py-2.5 rounded-xl border border-slate-300 focus:ring-2 focus:ring-indigo-500"
                                placeholder="Senha"
                                autoComplete="current-password"
                            />
                            <button type="submit" className="w-full bg-indigo-600 hover:bg-indigo-700 text-white py-2.5 rounded-xl font-bold transition-colors">Entrar</button>
                        </form>

                        <button onClick={() => switchMode('create')} className="mt-4 text-sm font-semibold text-indigo-600 hover:text-indigo-800 transition-colors" type="button">
                            Esqueci meus dados ou quero criar outro acesso
                        </button>
                    </div>
                )}

                {error && <p className="text-sm text-red-600 mt-4 font-semibold">{error}</p>}
            </div>
        </div>
    );
};

export default function App() {
    const [localDataRaw, setLocalData] = useLocalStorage('vagas_app_data', []);
    const localData = Array.isArray(localDataRaw) ? localDataRaw : [];

    const [history, setHistory] = useState({ past: [], present: localData, future: [] });
    const [customChartsRaw, setCustomCharts] = useLocalStorage('vagas_custom_charts', {});

    const [savedAccountRaw, setLegacySavedAccount] = useLocalStorage(AUTH_LEGACY_STORAGE_KEY, null);
    const [encryptedAccountDb, setEncryptedAccountDb] = useLocalStorage(AUTH_DB_STORAGE_KEY, '');
    const decryptedAccount = useMemo(() => readEncryptedExcelAuthDb(encryptedAccountDb), [encryptedAccountDb]);
    const legacyAccount = hasStoredAccount(savedAccountRaw) ? savedAccountRaw : null;
    const savedAccount = hasStoredAccount(decryptedAccount) ? decryptedAccount : legacyAccount;
    const [authSessionRaw, setAuthSession] = useLocalStorage('vagas_auth_session', null);
    const authSession = authSessionRaw && typeof authSessionRaw === 'object' ? authSessionRaw : null;
    const sessionMatchesSavedAccount = hasActiveAuthSession(authSession, savedAccount);
    const [isAuthenticated, setIsAuthenticated] = useState(sessionMatchesSavedAccount);
    const [currentUsername, setCurrentUsername] = useState(sessionMatchesSavedAccount ? savedAccount.username : '');
    const [authNotice, setAuthNotice] = useState('');

    const [showWelcome, setShowWelcome] = useState(false);
    const accountTutorialProgress = useMemo(() => normalizeTutorialProgress(savedAccount?.tutorialProgress || DEFAULT_TUTORIAL_PROGRESS), [savedAccount]);
    const [sessionTutorialProgress, setSessionTutorialProgress] = useState(DEFAULT_TUTORIAL_PROGRESS);
    const [showTutorial, setShowTutorial] = useState(false);
    const [showPatchNotes, setShowPatchNotes] = useState(false);
    const [tutorialSection, setTutorialSection] = useState('TABELA');
    const [pendingTutorialSection, setPendingTutorialSection] = useState(null);
    const [tutorialActiveStep, setTutorialActiveStep] = useState(null);

    const [loading, setLoading] = useState(false);
    const [isInitialSyncing, setIsInitialSyncing] = useState(false);
    const [activeTab, setActiveTab] = useState('TABELA');
    const [gridPage, setGridPage] = useState(1);
    const [searchTerm, setSearchTerm] = useState('');
    const [filters, setFilters] = useState({ status: 'TODOS', municipio: 'TODOS', motivo: 'TODOS', urgencia: 'TODOS' });
    const [showErrorsOnly, setShowErrorsOnly] = useState(false);
    const [sheetSearchTerm, setSheetSearchTerm] = useState('');
    const [sheetFilterColumn, setSheetFilterColumn] = useState('TODOS');
    const [sheetFilterTerm, setSheetFilterTerm] = useState('');
    const [sheetPage, setSheetPage] = useState(1);
    const [sheetUiPrefsRaw, setSheetUiPrefs] = useLocalStorage('vagas_sheets_ui_prefs', { hiddenColumns: [], widthOverrides: {} });
    const sheetUiPrefs = sheetUiPrefsRaw && typeof sheetUiPrefsRaw === 'object'
        ? sheetUiPrefsRaw
        : { hiddenColumns: [], widthOverrides: {} };
    const [isSheetColumnsPanelOpen, setIsSheetColumnsPanelOpen] = useState(false);
    const [isQuickAddModalOpen, setIsQuickAddModalOpen] = useState(false);
    const [quickAddData, setQuickAddData] = useState({});

    const [selectedRecord, setSelectedRecord] = useState(null);
    const [isEditing, setIsEditing] = useState(false);
    const [editFormData, setEditFormData] = useState({});
    const [isChartModalOpen, setIsChartModalOpen] = useState(false);
    const [newChartData, setNewChartData] = useState({ title: '', type: 'bar', groupBy: 'CARGO' });
    const [isGSheetsModalOpen, setIsGSheetsModalOpen] = useState(false);
    const [deleteConfirmOpen, setDeleteConfirmOpen] = useState(false);
    const [deleteModalRecord, setDeleteModalRecord] = useState(null);

    const [isCinematic, setIsCinematic] = useState(false);

    const mainScrollRef = useRef(null);
    const tableSearchInputRef = useRef(null);
    const sheetSearchInputRef = useRef(null);
    const sheetColumnsPanelRef = useRef(null);
    const sheetResizeStateRef = useRef(null);
    const persistAccount = (account) => {
        if (!hasStoredAccount(account)) {
            setEncryptedAccountDb('');
            setLegacySavedAccount(null);
            return;
        }

        const normalizedAccount = {
            ...account,
            tutorialProgress: normalizeTutorialProgress(account?.tutorialProgress || DEFAULT_TUTORIAL_PROGRESS),
        };

        const encryptedDb = createEncryptedExcelAuthDb(normalizedAccount);
        if (encryptedDb) {
            setEncryptedAccountDb(encryptedDb);
            setLegacySavedAccount(null);
            return;
        }

        // Fallback para nao perder login caso exista algum erro inesperado na serializacao.
        setLegacySavedAccount(normalizedAccount);
    };

    const scrollMainToTop = (behavior = 'smooth') => {
        if (mainScrollRef.current) mainScrollRef.current.scrollTo({ top: 0, behavior });
    };

    const data = Array.isArray(history.present) ? history.present : [];
    const chartsOwnerKey = useMemo(() => normalizeCredentialText(savedAccount?.username || currentUsername || 'global') || 'global', [currentUsername, savedAccount]);
    const customChartsStore = useMemo(() => {
        if (Array.isArray(customChartsRaw)) return { global: customChartsRaw };
        if (customChartsRaw && typeof customChartsRaw === 'object') return customChartsRaw;
        return {};
    }, [customChartsRaw]);
    const customCharts = useMemo(() => {
        const byUser = customChartsStore[chartsOwnerKey];
        if (Array.isArray(byUser)) return byUser;
        if (Array.isArray(customChartsStore.global)) return customChartsStore.global;
        return [];
    }, [chartsOwnerKey, customChartsStore]);
    const tutorialSteps = TUTORIAL_STEPS[tutorialSection] || [];
    const hasCompletedTutorialSection = (section) => accountTutorialProgress[section] || sessionTutorialProgress[section];
    const resetAuthenticatedUi = () => {
        setIsAuthenticated(false);
        setCurrentUsername('');
        setShowWelcome(false);
        setShowTutorial(false);
        setPendingTutorialSection(null);
        setTutorialActiveStep(null);
        setSelectedRecord(null);
        setIsEditing(false);
        setEditFormData({});
        setIsChartModalOpen(false);
        setIsGSheetsModalOpen(false);
    };

    useEffect(() => {
        if (!legacyAccount || hasStoredAccount(decryptedAccount)) return;
        const encryptedDb = createEncryptedExcelAuthDb(legacyAccount);
        if (encryptedDb) {
            setEncryptedAccountDb(encryptedDb);
            setLegacySavedAccount(null);
        }
    }, [decryptedAccount, legacyAccount, setEncryptedAccountDb, setLegacySavedAccount]);

    useEffect(() => {
        if (!Array.isArray(customChartsRaw)) return;
        setCustomCharts((previous) => {
            if (!Array.isArray(previous)) return previous;
            return {
                [chartsOwnerKey]: previous,
                global: previous,
            };
        });
    }, [chartsOwnerKey, customChartsRaw, setCustomCharts]);

    useEffect(() => {
        if (!savedAccount) {
            if (authSession) setAuthSession(null);
            setIsAuthenticated(false);
            setCurrentUsername('');
            return;
        }

        if (!authSession?.authenticated || !matchesSavedUsername(authSession.username, savedAccount.username)) {
            setIsAuthenticated(false);
            setCurrentUsername('');
            return undefined;
        }

        const expiryTimestamp = getAuthSessionExpiryTimestamp(authSession);
        if (!expiryTimestamp || expiryTimestamp <= Date.now()) {
            setAuthSession(null);
            setAuthNotice('Sua sessao expirou apos 24 horas. Faca login novamente.');
            resetAuthenticatedUi();
            return undefined;
        }

        setCurrentUsername(savedAccount.username);
        setIsAuthenticated(true);

        const timer = window.setTimeout(() => {
            setAuthSession(null);
            setAuthNotice('Sua sessao expirou apos 24 horas. Faca login novamente.');
            resetAuthenticatedUi();
        }, expiryTimestamp - Date.now());

        return () => window.clearTimeout(timer);
    }, [authSession, savedAccount]);

    useEffect(() => {
        if (showWelcome || !isAuthenticated || data.length === 0 || showTutorial || activeTab !== 'TABELA' || hasCompletedTutorialSection('TABELA')) {
            return undefined;
        }

        const timer = setTimeout(() => {
            setTutorialSection('TABELA');
            setShowTutorial(true);
        }, 250);

        return () => clearTimeout(timer);
    }, [activeTab, data.length, isAuthenticated, accountTutorialProgress, sessionTutorialProgress, showTutorial, showWelcome]);

    useEffect(() => {
        if (showWelcome || !isAuthenticated || data.length === 0 || showTutorial || showPatchNotes || activeTab !== 'TABELA' || hasCompletedTutorialSection('NEW_FEATURES_TABLE')) {
            return undefined;
        }

        if (!hasCompletedTutorialSection('PATCH_NOTES_VIEWED')) {
            const timer = setTimeout(() => {
                setShowPatchNotes(true);
            }, 350);
            return () => clearTimeout(timer);
        }

        const timer = setTimeout(() => {
            setTutorialSection('NEW_FEATURES_TABLE');
            setShowTutorial(true);
        }, 350);

        return () => clearTimeout(timer);
    }, [activeTab, data.length, isAuthenticated, accountTutorialProgress, sessionTutorialProgress, showTutorial, showWelcome, showPatchNotes]);

    useEffect(() => {
        if (showWelcome || !isAuthenticated || data.length === 0 || showTutorial || activeTab !== 'SHEETS' || hasCompletedTutorialSection('NEW_FEATURES_SHEETS_V2')) {
            return undefined;
        }

        const timer = setTimeout(() => {
            setTutorialSection('NEW_FEATURES_SHEETS_V2');
            setShowTutorial(true);
        }, 350);

        return () => clearTimeout(timer);
    }, [activeTab, data.length, isAuthenticated, accountTutorialProgress, sessionTutorialProgress, showTutorial, showWelcome]);

    useEffect(() => {
        if (!pendingTutorialSection || activeTab !== pendingTutorialSection || showTutorial || showWelcome) {
            return undefined;
        }

        const timer = setTimeout(() => {
            setTutorialSection(pendingTutorialSection);
            setShowTutorial(true);
            setPendingTutorialSection(null);
        }, 250);

        return () => clearTimeout(timer);
    }, [activeTab, pendingTutorialSection, showTutorial, showWelcome]);

    const handlePatchNotesClose = () => {
        setShowPatchNotes(false);
        if (savedAccount) {
            persistAccount({
                ...savedAccount,
                tutorialProgress: {
                    ...normalizeTutorialProgress(savedAccount.tutorialProgress || DEFAULT_TUTORIAL_PROGRESS),
                    PATCH_NOTES_VIEWED: true,
                },
                updatedAt: new Date().toISOString(),
            });
        }
        setSessionTutorialProgress((previous) => ({
            ...previous,
            PATCH_NOTES_VIEWED: true,
        }));
    };

    const handlePatchNotesStartTableTutorial = () => {
        handlePatchNotesClose();
        setTimeout(() => {
            setActiveTab('TABELA');
            setTutorialSection('NEW_FEATURES_TABLE');
            setShowTutorial(true);
        }, 200);
    };

    const handlePatchNotesStartSheetsTutorial = () => {
        handlePatchNotesClose();
        setTimeout(() => {
            setActiveTab('SHEETS');
            setTutorialSection('NEW_FEATURES_SHEETS_V2');
            setShowTutorial(true);
        }, 220);
    };

    const handleTutorialComplete = (section, dontShow) => {
        if (dontShow && savedAccount) {
            persistAccount({
                ...savedAccount,
                tutorialProgress: {
                    ...normalizeTutorialProgress(savedAccount.tutorialProgress || DEFAULT_TUTORIAL_PROGRESS),
                    [section]: true,
                },
                updatedAt: new Date().toISOString(),
            });
        }

        setSessionTutorialProgress((previous) => ({
            ...previous,
            [section]: true,
        }));

        setPendingTutorialSection(null);
        setShowTutorial(false);
        setTutorialActiveStep(null);
        setSelectedRecord(null);
        setIsEditing(false);
        setEditFormData({});
        setIsChartModalOpen(false);
        setIsQuickAddModalOpen(false);
        if (typeof window !== 'undefined') window.requestAnimationFrame(() => scrollMainToTop());
    };

    const handleCreateAccount = (account) => {
        persistAccount(account);
        setAuthSession({ authenticated: true, username: account.username, at: new Date().toISOString() });
        setAuthNotice('');
        setCurrentUsername(account.username);
        setIsAuthenticated(true);
        setShowWelcome(true);
    };

    const handleLoginSuccess = () => {
        setAuthSession({ authenticated: true, username: savedAccount?.username || '', at: new Date().toISOString() });
        setAuthNotice('');
        setCurrentUsername(savedAccount?.username || '');
        setIsAuthenticated(true);
        setShowWelcome(true);
    };

    const handleLogout = () => {
        setAuthSession(null);
        setAuthNotice('');
        resetAuthenticatedUi();
    };

    const processDataImport = (newDataArray, fromSheets, isSilent = false) => {
        const cleanedNewData = newDataArray.filter((row) => Object.values(row).some((val) => val !== null && String(val).trim() !== ''));
        setAppData((prevData) => {
            const prevArray = Array.isArray(prevData) ? prevData : [];
            if (prevArray.length === 0) return validateData(cleanedNewData.map((row) => ({ ...row, _id: Math.random().toString(36).slice(2, 11) })));
            const mergedData = [...prevArray];
            cleanedNewData.forEach((newRow) => {
                const existingIdx = mergedData.findIndex((r) => (r['Mat. Subs'] && r['Mat. Subs'] === newRow['Mat. Subs']) || (r['Nº Protoc'] && r['Nº Protoc'] === newRow['Nº Protoc']));
                if (existingIdx >= 0) {
                    const preserve = {
                        Candidato: mergedData[existingIdx].Candidato !== 'SEM COBERTURA' ? mergedData[existingIdx].Candidato : newRow.Candidato,
                        Status: mergedData[existingIdx].Status !== 'ABERTA' ? mergedData[existingIdx].Status : newRow.Status,
                        'Contato Candidato': mergedData[existingIdx]['Contato Candidato'] || newRow['Contato Candidato'],
                        'OBS:': mergedData[existingIdx]['OBS:'] || newRow['OBS:'],
                    };
                    mergedData[existingIdx] = { ...mergedData[existingIdx], ...newRow, ...preserve };
                } else mergedData.push({ ...newRow, _id: Math.random().toString(36).slice(2, 11) });
            });
            return validateData(mergedData);
        });
        setLoading(false);
    };

    useEffect(() => {
        const autoSyncWithSheets = async () => {
            if (localData.length === 0) {
                setIsInitialSyncing(true);
                try {
                    const response = await fetch(GOOGLE_SHEETS_CSV_EXPORT);
                    if (response.ok) {
                        const csvText = await response.text();
                        const parsed = parseCSV(csvText);
                        if (parsed && parsed.length > 0) processDataImport(parsed, true, true);
                    }
                } catch (error) {
                    // no-op
                }
                setIsInitialSyncing(false);
            }
        };
        autoSyncWithSheets();
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);

    const setAppData = (action) => {
        const newPresent = typeof action === 'function' ? action(history.present) : action;
        if (JSON.stringify(newPresent) === JSON.stringify(history.present)) return;
        setHistory((h) => ({ past: [...h.past, h.present].slice(-30), present: newPresent, future: [] }));
        setLocalData(newPresent);
    };

    const handleUndo = () => {
        setHistory((h) => {
            if (h.past.length === 0) return h;
            const previous = h.past[h.past.length - 1];
            const newPast = h.past.slice(0, h.past.length - 1);
            setLocalData(previous);
            return { past: newPast, present: previous, future: [h.present, ...h.future] };
        });
    };

    const handleRedo = () => {
        setHistory((h) => {
            if (h.future.length === 0) return h;
            const next = h.future[0];
            const newFuture = h.future.slice(1);
            setLocalData(next);
            return { past: [...h.past, h.present], present: next, future: newFuture };
        });
    };

    const triggerCinematicTransition = (callback) => {
        setIsCinematic(true);
        setTimeout(() => {
            callback();
            scrollMainToTop('auto');
            setTimeout(() => setIsCinematic(false), 50);
        }, 350);
    };

    const handleTabChange = (newTab, options = {}) => {
        const { openTutorial = false } = options;
        const shouldQueueTutorial = openTutorial
            && isAuthenticated
            && !showWelcome
            && !showTutorial
            && data.length > 0
            && !hasCompletedTutorialSection(newTab);

        if (activeTab === newTab) {
            if (shouldQueueTutorial) {
                setTutorialSection(newTab);
                setShowTutorial(true);
            }
            return;
        }

        triggerCinematicTransition(() => {
            setActiveTab(newTab);
            if (shouldQueueTutorial) setPendingTutorialSection(newTab);
        });
    };


    const validateData = (rows) => {
        const safeRows = Array.isArray(rows) ? rows : [];

        return safeRows
            .filter((row) => String(getRowValue(row, ['Nome Subs'])).trim() !== '')
            .map((row) => ({
                ...row,
                _isInvalid: String(getRowValue(row, ['Status'])).trim() === '',
            }));
    };

    const handleGoogleSheetsSync = async () => {
        setIsGSheetsModalOpen(false);
        setLoading(true);
        try {
            const response = await fetch(GOOGLE_SHEETS_CSV_EXPORT);
            if (!response.ok) throw new Error('CORS');
            processDataImport(parseCSV(await response.text()), true);
        } catch (error) {
            setIsGSheetsModalOpen(true);
            setLoading(false);
        }
    };

    const handleFileUpload = async (event) => {
        const file = event.target.files[0];
        if (!file) return;
        setLoading(true);
        setIsGSheetsModalOpen(false);
        try {
            const buffer = await file.arrayBuffer();
            const workbook = XLSX.read(buffer, { type: 'array' });
            const worksheet = workbook.Sheets.PAINEL || workbook.Sheets[workbook.SheetNames[0]];
            processDataImport(XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false }), false);
        } catch (error) {
            alert('Erro de parsing.');
        } finally {
            setLoading(false);
            event.target.value = null;
        }
    };

    const handleExportExcel = () => {
        if (data.length === 0) return;
        const exportData = data.map((row) => {
            const newRow = { ...row };
            delete newRow._id;
            delete newRow._isInvalid;
            return newRow;
        });
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(exportData), 'PAINEL_ATUALIZADO');
        XLSX.writeFile(wb, `Vagas_${new Date().toISOString().slice(0, 10)}.xlsx`);
    };

    const handleInlineEdit = (id, field, value) => {
        setAppData((prevData) => validateData((Array.isArray(prevData) ? prevData : []).map((item) => (item._id === id ? { ...item, [field]: value } : item))));
    };

    const openQuickAddModal = () => {
        const base = {};
        detailedColumns.forEach((column) => {
            base[column] = '';
        });
        if (Object.prototype.hasOwnProperty.call(base, 'Status')) base.Status = 'ABERTA';
        setQuickAddData(base);
        setIsQuickAddModalOpen(true);
    };

    const handleQuickAddSave = () => {
        const requiredMissing = REQUIRED_FIELDS.filter((field) => {
            const key = getRowKey(quickAddData, [field]);
            return !String(quickAddData[key] || '').trim();
        });

        if (requiredMissing.length > 0) {
            alert(`Preencha os campos obrigatorios: ${requiredMissing.join(', ')}`);
            return;
        }

        const statusKey = getRowKey(quickAddData, ['Status']);

        const record = {
            ...quickAddData,
            [statusKey || 'Status']: String(quickAddData[statusKey] || 'ABERTA').trim() || 'ABERTA',
            _id: Date.now(),
        };

        setAppData((prev) => validateData([record, ...(Array.isArray(prev) ? prev : [])]));
        setIsQuickAddModalOpen(false);
    };

    const handleEditClick = () => {
        setDeleteConfirmOpen(false);
        const normalized = { ...selectedRecord };
        detailedColumns.forEach((column) => {
            if (normalized[column] === undefined || normalized[column] === null) normalized[column] = '';
        });
        setEditFormData(normalized);
        setIsEditing(true);
    };

    const handleSaveEdit = () => {
        setAppData((prev) => validateData((Array.isArray(prev) ? prev : []).map((item) => (item._id === editFormData._id ? editFormData : item))));
        setSelectedRecord(editFormData);
        setIsEditing(false);
    };

    const handleDeleteRecord = () => {
        setAppData((prev) => validateData((Array.isArray(prev) ? prev : []).filter((item) => item._id !== selectedRecord._id)));
        setSelectedRecord(null);
        setDeleteConfirmOpen(false);
        setIsEditing(false);
        setEditFormData({});
    };

    const requestDeleteRecord = (record) => {
        if (!record?._id) return;
        setDeleteModalRecord(record);
    };

    const handleDeleteModalConfirm = () => {
        const recordId = deleteModalRecord?._id;
        if (!recordId) return;

        setAppData((prev) => validateData((Array.isArray(prev) ? prev : []).filter((item) => item._id !== recordId)));
        if (selectedRecord?._id === recordId) {
            setSelectedRecord(null);
            setDeleteConfirmOpen(false);
            setIsEditing(false);
            setEditFormData({});
        }
        setDeleteModalRecord(null);
    };

    const handleDeleteModalCancel = () => {
        setDeleteModalRecord(null);
    };

    const addCustomChart = () => {
        if (newChartData.title && newChartData.groupBy) {
            setCustomCharts((previous) => {
                const currentStore = Array.isArray(previous)
                    ? { global: previous }
                    : (previous && typeof previous === 'object' ? previous : {});

                const currentUserCharts = Array.isArray(currentStore[chartsOwnerKey]) ? currentStore[chartsOwnerKey] : [];

                return {
                    ...currentStore,
                    [chartsOwnerKey]: [...currentUserCharts, { id: Date.now(), ...newChartData }],
                };
            });
            setIsChartModalOpen(false);
            setNewChartData({ title: '', type: 'bar', groupBy: 'CARGO' });
        }
    };

    const removeCustomChart = (chartId) => {
        setCustomCharts((previous) => {
            const currentStore = Array.isArray(previous)
                ? { global: previous }
                : (previous && typeof previous === 'object' ? previous : {});

            const currentUserCharts = Array.isArray(currentStore[chartsOwnerKey]) ? currentStore[chartsOwnerKey] : [];

            return {
                ...currentStore,
                [chartsOwnerKey]: currentUserCharts.filter((chart) => chart.id !== chartId),
            };
        });
    };

    const handleChartClick = (groupBy, value) => {
        if (!value || value === 'NAO INFORMADO') return;
        setFilters({ status: 'TODOS', municipio: 'TODOS', motivo: 'TODOS', urgencia: 'TODOS' });
        setSearchTerm('');
        setShowErrorsOnly(false);
        if (groupBy === 'Status') setFilters((prev) => ({ ...prev, status: value.toUpperCase() }));
        else if (groupBy === 'NRE / MUNICIPIO') setFilters((prev) => ({ ...prev, municipio: value }));
        else if (groupBy === 'Motivo') setFilters((prev) => ({ ...prev, motivo: value }));
        else setSearchTerm(value);
        triggerCinematicTransition(() => setActiveTab('TABELA'));
    };

    const availableColumns = useMemo(
        () => (data.length === 0 ? [] : Object.keys(data[0]).filter((k) => !['_id', '_isInvalid'].includes(k) && data[0][k] !== undefined)),
        [data],
    );

    const listOptions = useMemo(() => {
        const m = new Set();
        const mt = new Set();
        const s = new Set();
        data.forEach((i) => {
            if (i['NRE / MUNICIPIO']) m.add(String(i['NRE / MUNICIPIO']).trim());
            if (i.Motivo) mt.add(String(i.Motivo).trim());
            if (i.Status) s.add(String(i.Status).trim().toUpperCase());
        });
        return { municipios: Array.from(m).sort(), motivos: Array.from(mt).sort(), status: Array.from(s).sort() };
    }, [data]);

    const metrics = useMemo(() => {
        if (data.length === 0) return null;
        const invalidCount = data.filter((d) => d._isInvalid).length;
        const generateChartData = (groupBy, limit = 10) => Object.entries(
            data.reduce((acc, curr) => {
                const v = curr[groupBy] ? String(curr[groupBy]).trim() : 'NAO INFORMADO';
                if (v && v !== 'NAO INFORMADO') acc[v] = (acc[v] || 0) + 1;
                return acc;
            }, {}),
        )
            .map(([k, v]) => ({ name: k, value: v }))
            .sort((a, b) => b.value - a.value)
            .slice(0, limit);

        return {
            total: data.length,
            invalidCount,
            abertas: data.filter((d) => safeGet(d, 'Status') === 'ABERTA').length,
            fechadas: data.filter((d) => safeGet(d, 'Status') === 'FECHADA').length,
            semCobertura: data.filter((d) => safeGet(d, 'Candidato') === 'SEM COBERTURA' || !d.Candidato).length,
            statusChartData: generateChartData('Status'),
            motivoChartData: generateChartData('Motivo', 5),
            generateChartData,
        };
    }, [data]);

    const filteredData = useMemo(() => {
        const filtered = data.filter((item) => {
            if (showErrorsOnly) return item._isInvalid;
            const mSearch = Object.values(item).some((val) => String(val).toLowerCase().includes(searchTerm.toLowerCase()));
            const mStatus = filters.status === 'TODOS' || safeGet(item, 'Status') === filters.status;
            const mMun = filters.municipio === 'TODOS' || String(item['NRE / MUNICIPIO'] || '').trim() === filters.municipio;
            let mUrg = true;
            if (filters.urgencia !== 'TODOS') {
                const dias = getDaysDiff(getPrazoValue(item));
                if (filters.urgencia === 'URGENTE') mUrg = dias !== null && dias <= 5;
                if (filters.urgencia === 'MEDIA') mUrg = dias === null || (dias > 5 && dias <= 30);
                if (filters.urgencia === 'LONGE') mUrg = dias !== null && dias > 30;
            }
            return mSearch && mStatus && mMun && mUrg && (filters.motivo === 'TODOS' || String(item.Motivo || '').trim() === filters.motivo);
        });

        return filtered.sort((a, b) => {
            if (a._isInvalid && !b._isInvalid) return -1;
            if (!a._isInvalid && b._isInvalid) return 1;
            const isResA = ['FECHADA', 'ENCAMINHADA', 'CANCELADA'].includes(safeGet(a, 'Status')) || (a.Candidato && a.Candidato !== 'SEM COBERTURA');
            const isResB = ['FECHADA', 'ENCAMINHADA', 'CANCELADA'].includes(safeGet(b, 'Status')) || (b.Candidato && b.Candidato !== 'SEM COBERTURA');
            if (!isResA && isResB) return -1;
            if (isResA && !isResB) return 1;
            if (!isResA && !isResB) {
                const d_A = getDaysDiff(getPrazoValue(a));
                const d_B = getDaysDiff(getPrazoValue(b));
                if (d_A !== null && d_B !== null) return d_A - d_B;
                if (d_A !== null && d_B === null) return -1;
                if (d_A === null && d_B !== null) return 1;
            }
            return 0;
        });
    }, [data, searchTerm, filters, showErrorsOnly]);

    const sheetsColumns = useMemo(() => {
        const allColumns = Array.from(new Set(filteredData.flatMap((row) => Object.keys(row || {}))))
            .filter((column) => !['_id', '_isInvalid'].includes(column));

        const prioritized = SHEETS_PRIORITY_COLUMNS.filter((column) => allColumns.includes(column));
        const others = allColumns.filter((column) => !SHEETS_PRIORITY_COLUMNS.includes(column)).sort((a, b) => a.localeCompare(b));

        return [...prioritized, ...others];
    }, [filteredData]);

    const detailedColumns = useMemo(() => {
        if (sheetsColumns.length > 0) return sheetsColumns;
        if (selectedRecord && typeof selectedRecord === 'object') {
            return Object.keys(selectedRecord).filter((column) => !['_id', '_isInvalid'].includes(column));
        }
        return SHEETS_PRIORITY_COLUMNS;
    }, [selectedRecord, sheetsColumns]);

    const sheetFilteredData = useMemo(() => {
        let result = filteredData;

        // Aplicar busca por texto em todas as colunas
        if (sheetSearchTerm.trim()) {
            const searchLower = sheetSearchTerm.toLowerCase();
            result = result.filter(row =>
                Object.values(row).some(val =>
                    String(val || '').toLowerCase().includes(searchLower)
                )
            );
        }

        // Aplicar filtro simples por coluna selecionada
        const normalizedTerm = String(sheetFilterTerm || '').trim().toLowerCase();
        if (sheetFilterColumn !== 'TODOS' && normalizedTerm) {
            result = result.filter((row) => String(row[sheetFilterColumn] || '').toLowerCase().includes(normalizedTerm));
        }

        return result;
    }, [filteredData, sheetFilterColumn, sheetFilterTerm, sheetSearchTerm]);

    const hiddenSheetColumns = useMemo(() => (
        Array.isArray(sheetUiPrefs.hiddenColumns) ? sheetUiPrefs.hiddenColumns : []
    ), [sheetUiPrefs.hiddenColumns]);

    const visibleSheetColumns = useMemo(() => (
        sheetsColumns.filter((column) => !hiddenSheetColumns.includes(column))
    ), [hiddenSheetColumns, sheetsColumns]);

    const sheetAutoColumnWidths = useMemo(() => {
        const targetColumns = visibleSheetColumns;
        if (targetColumns.length === 0) return {};

        const sampleRows = sheetFilteredData.slice(0, 140);
        const clamp = (value, min, max) => Math.max(min, Math.min(max, value));

        return targetColumns.reduce((acc, column) => {
            const normalized = normalizeCredentialText(column);
            const headerLen = String(column || '').trim().length;
            const lengths = sampleRows
                .map((row) => String(row?.[column] || '').trim().length)
                .filter((len) => len > 0)
                .sort((a, b) => a - b);

            const avgLen = lengths.length > 0 ? (lengths.reduce((sum, len) => sum + len, 0) / lengths.length) : 0;
            const p75Len = lengths.length > 0 ? lengths[Math.floor((lengths.length - 1) * 0.75)] : 0;
            let px = Math.max(headerLen * 8, avgLen * 7.2, p75Len * 6.5, 88);

            if (normalized.includes('nome subs')) px += 65;
            if (normalized === 'cargo') px += 70;
            if (normalized === 'candidato') px += 40;
            if (normalized.includes('contato candidato')) px += 35;
            if (normalized.includes('municipio') || normalized.includes('nre')) px += 45;
            if (normalized === 'motivo') px += 45;
            if (normalized.includes('obs')) px += 95;

            acc[column] = clamp(Math.round(px), 70, 340);
            return acc;
        }, {});
    }, [sheetFilteredData, visibleSheetColumns]);

    const sheetWidthOverrides = useMemo(() => (
        sheetUiPrefs.widthOverrides && typeof sheetUiPrefs.widthOverrides === 'object'
            ? sheetUiPrefs.widthOverrides
            : {}
    ), [sheetUiPrefs.widthOverrides]);

    const sheetColumnWidths = useMemo(() => (
        visibleSheetColumns.reduce((acc, column) => {
            const overrideWidth = Number(sheetWidthOverrides[column]);
            acc[column] = Number.isFinite(overrideWidth) && overrideWidth > 40
                ? overrideWidth
                : (sheetAutoColumnWidths[column] || 100);
            return acc;
        }, {})
    ), [sheetAutoColumnWidths, sheetWidthOverrides, visibleSheetColumns]);

    const gridTotalPages = useMemo(() => Math.max(1, Math.ceil(filteredData.length / GRID_PAGE_SIZE)), [filteredData.length]);
    const gridPagedData = useMemo(() => {
        const safePage = Math.min(gridPage, gridTotalPages);
        const start = (safePage - 1) * GRID_PAGE_SIZE;
        return filteredData.slice(start, start + GRID_PAGE_SIZE);
    }, [filteredData, gridPage, gridTotalPages]);

    const sheetTotalPages = useMemo(() => Math.max(1, Math.ceil(sheetFilteredData.length / SHEETS_PAGE_SIZE)), [sheetFilteredData.length]);
    const sheetPagedData = useMemo(() => {
        const safePage = Math.min(sheetPage, sheetTotalPages);
        const start = (safePage - 1) * SHEETS_PAGE_SIZE;
        return sheetFilteredData.slice(start, start + SHEETS_PAGE_SIZE);
    }, [sheetFilteredData, sheetPage, sheetTotalPages]);

    useEffect(() => {
        setGridPage(1);
    }, [searchTerm, filters, showErrorsOnly]);

    useEffect(() => {
        setSheetPage(1);
    }, [sheetFilterColumn, sheetFilterTerm, sheetSearchTerm]);

    useEffect(() => {
        if (gridPage > gridTotalPages) setGridPage(gridTotalPages);
    }, [gridPage, gridTotalPages]);

    useEffect(() => {
        if (sheetPage > sheetTotalPages) setSheetPage(sheetTotalPages);
    }, [sheetPage, sheetTotalPages]);

    useEffect(() => {
        setSheetUiPrefs((current) => {
            const safe = current && typeof current === 'object' ? current : { hiddenColumns: [], widthOverrides: {} };
            const hidden = Array.isArray(safe.hiddenColumns)
                ? safe.hiddenColumns.filter((column) => sheetsColumns.includes(column))
                : [];
            const widthOverrides = safe.widthOverrides && typeof safe.widthOverrides === 'object'
                ? Object.fromEntries(Object.entries(safe.widthOverrides).filter(([column]) => sheetsColumns.includes(column)))
                : {};

            if (
                JSON.stringify(hidden) === JSON.stringify(safe.hiddenColumns || [])
                && JSON.stringify(widthOverrides) === JSON.stringify(safe.widthOverrides || {})
            ) {
                return safe;
            }

            return { hiddenColumns: hidden, widthOverrides };
        });
    }, [setSheetUiPrefs, sheetsColumns]);

    useEffect(() => {
        const handleOutsideClick = (event) => {
            if (!isSheetColumnsPanelOpen) return;
            if (sheetColumnsPanelRef.current && !sheetColumnsPanelRef.current.contains(event.target)) {
                setIsSheetColumnsPanelOpen(false);
            }
        };

        document.addEventListener('mousedown', handleOutsideClick);
        return () => document.removeEventListener('mousedown', handleOutsideClick);
    }, [isSheetColumnsPanelOpen]);

    useEffect(() => {
        const handleMouseMove = (event) => {
            const state = sheetResizeStateRef.current;
            if (!state) return;
            const delta = event.clientX - state.startX;
            const nextWidth = Math.max(70, Math.min(420, Math.round(state.startWidth + delta)));

            setSheetUiPrefs((current) => {
                const safe = current && typeof current === 'object' ? current : { hiddenColumns: [], widthOverrides: {} };
                const currentOverrides = safe.widthOverrides && typeof safe.widthOverrides === 'object' ? safe.widthOverrides : {};
                return {
                    ...safe,
                    widthOverrides: {
                        ...currentOverrides,
                        [state.column]: nextWidth,
                    },
                };
            });
        };

        const handleMouseUp = () => {
            if (sheetResizeStateRef.current) sheetResizeStateRef.current = null;
        };

        window.addEventListener('mousemove', handleMouseMove);
        window.addEventListener('mouseup', handleMouseUp);
        return () => {
            window.removeEventListener('mousemove', handleMouseMove);
            window.removeEventListener('mouseup', handleMouseUp);
        };
    }, [setSheetUiPrefs]);

    const toggleSheetColumnVisibility = (column) => {
        setSheetUiPrefs((current) => {
            const safe = current && typeof current === 'object' ? current : { hiddenColumns: [], widthOverrides: {} };
            const hidden = Array.isArray(safe.hiddenColumns) ? safe.hiddenColumns : [];
            const isHidden = hidden.includes(column);
            const nextHidden = isHidden ? hidden.filter((item) => item !== column) : [...hidden, column];

            if (nextHidden.length >= sheetsColumns.length) return safe;

            return { ...safe, hiddenColumns: nextHidden };
        });
    };

    const resetSheetColumnsLayout = () => {
        setSheetUiPrefs({ hiddenColumns: [], widthOverrides: {} });
    };

    const startSheetColumnResize = (column, event) => {
        event.preventDefault();
        event.stopPropagation();
        const startWidth = sheetColumnWidths[column] || 120;
        sheetResizeStateRef.current = {
            column,
            startX: event.clientX,
            startWidth,
        };
    };

    useEffect(() => {
        if (!showTutorial || tutorialSection !== 'TABELA') return;

        const currentStepId = tutorialActiveStep?.id;
        if (!TUTORIAL_RECORD_MODAL_STEP_IDS.has(currentStepId)) {
            setSelectedRecord(null);
            setIsEditing(false);
            return;
        }

        const previewRow = filteredData[0];
        if (!previewRow) return;

        setSelectedRecord((current) => (current?._id === previewRow._id ? current : previewRow));

        if (TUTORIAL_RECORD_EDIT_STEP_IDS.has(currentStepId)) {
            setEditFormData((current) => (current?._id === previewRow._id ? current : { ...previewRow }));
            setIsEditing(true);
        } else {
            setIsEditing(false);
        }
    }, [filteredData, showTutorial, tutorialActiveStep, tutorialSection]);

    useEffect(() => {
        if (!showTutorial || tutorialSection !== 'DASHBOARD') {
            setIsChartModalOpen(false);
            return;
        }

        const currentStepId = tutorialActiveStep?.id;
        setIsChartModalOpen(TUTORIAL_CHART_MODAL_STEP_IDS.has(currentStepId));
    }, [showTutorial, tutorialActiveStep, tutorialSection]);

    useEffect(() => {
        if (!showTutorial || tutorialSection !== 'NEW_FEATURES_TABLE') return;

        const currentStepId = tutorialActiveStep?.id;
        setIsQuickAddModalOpen(currentStepId === 'new-features-quick-add-modal');
    }, [showTutorial, tutorialActiveStep, tutorialSection]);

    useEffect(() => {
        if (!showTutorial || tutorialSection !== 'NEW_FEATURES_SHEETS_V2') return;

        const currentStepId = tutorialActiveStep?.id;

        if (currentStepId === 'new-features-columns-panel' || currentStepId === 'new-features-columns-toggle') {
            setIsSheetColumnsPanelOpen(true);
        }

        if (currentStepId === 'new-features-actions-delete' || currentStepId === 'new-features-horizontal-scroll') {
            const scroller = document.getElementById('tour-sheets-horizontal-scroll');
            if (scroller) {
                scroller.scrollTo({ left: scroller.scrollWidth, behavior: 'smooth' });
            }
        }
    }, [showTutorial, tutorialActiveStep, tutorialSection]);

    useEffect(() => {
        const onKeyDown = (event) => {
            if (!isAuthenticated) return;
            if (event.ctrlKey || event.metaKey || event.altKey) return;

            const targetTag = String(event.target?.tagName || '').toLowerCase();
            const isTypingContext = ['input', 'textarea', 'select'].includes(targetTag) || event.target?.isContentEditable;

            if (event.key.toLowerCase() === 'c' && !isTypingContext && !showTutorial) {
                event.preventDefault();
                openQuickAddModal();
            }

            if (event.key.toLowerCase() === 'f' && !isTypingContext) {
                event.preventDefault();
                const inputToFocus = activeTab === 'SHEETS' ? sheetSearchInputRef.current : tableSearchInputRef.current;
                if (inputToFocus) {
                    inputToFocus.focus();
                    inputToFocus.select?.();
                }
            }
        };

        window.addEventListener('keydown', onKeyDown);
        return () => window.removeEventListener('keydown', onKeyDown);
    }, [activeTab, isAuthenticated, showTutorial]);

    const tableAnimationKey = `${activeTab}-${searchTerm}-${filters.status}-${filters.municipio}-${filters.urgencia}-${showErrorsOnly}`;

    return (
        <React.Fragment>
            {!isAuthenticated && <LoginOverlay account={savedAccount} onCreateAccount={handleCreateAccount} onLogin={handleLoginSuccess} notice={authNotice} />}
            {showWelcome && isAuthenticated && <WelcomeOverlay userName={currentUsername} onFinish={() => setShowWelcome(false)} />}

            <div className={`h-screen w-full bg-slate-50 flex flex-col relative overflow-hidden transition-opacity duration-500 ease-in-out ${(showWelcome || !isAuthenticated) ? 'opacity-0' : 'opacity-100'}`}>
                <style>{`
          @keyframes cascadeSlide { 0% { opacity: 0; transform: translateY(15px); } 100% { opacity: 1; transform: translateY(0); } }
          .anim-cascade { animation: cascadeSlide 0.35s ease-out forwards; opacity: 0; }
          .cinematic-effect { filter: blur(8px) grayscale(10%) brightness(0.9); transform: scale(0.98); opacity: 0.5; pointer-events: none; }
        `}</style>

                <header className="bg-white border-b border-slate-200 z-20 shadow-sm shrink-0">
                    <div className={`${activeTab === 'SHEETS' ? 'max-w-[2200px]' : 'max-w-[1600px]'} mx-auto px-4 sm:px-6 lg:px-8 py-3`}>
                        <div className="grid grid-cols-1 xl:grid-cols-[auto_minmax(320px,1fr)_minmax(520px,1.2fr)] items-center gap-3 xl:gap-5">
                            <div className="flex items-center gap-3 min-w-0 shrink-0">
                                <div className="bg-indigo-600 p-2.5 rounded-xl shadow-sm shrink-0"><Briefcase className="w-5 h-5 text-white" /></div>
                                <div className="min-w-0">
                                    <h1 className="text-xl font-black text-slate-800 tracking-tight whitespace-nowrap">Recrutamento</h1>
                                    <p className="hidden 2xl:block text-[11px] font-bold uppercase tracking-[0.22em] text-slate-400">Controle de admissoes</p>
                                </div>
                            </div>

                            <div className="min-w-0">
                                <div id="tour-tabs" className="flex items-center gap-1.5 bg-slate-100 p-1 rounded-xl overflow-x-auto hide-scrollbar w-full shadow-inner shadow-slate-200/60">
                                    <button onClick={() => handleTabChange('TABELA')} className={`flex-1 inline-flex items-center justify-center gap-2 min-w-[148px] px-4 py-2.5 rounded-lg text-sm font-semibold transition-all duration-300 whitespace-nowrap ${activeTab === 'TABELA' ? 'bg-white text-indigo-700 shadow-sm' : 'text-slate-500 hover:bg-white/60 hover:text-slate-700'}`} type="button"><ListTodo className="w-4 h-4" /> <span>Tabela</span>{metrics?.invalidCount > 0 && <span className="inline-flex items-center gap-1.5 ml-1"><span className="w-2 h-2 rounded-full bg-red-500 animate-pulse" /><span className="text-[11px] font-bold text-red-600">{metrics.invalidCount}</span></span>}</button>
                                    <button id="tour-tab-dashboard" onClick={() => handleTabChange('DASHBOARD', { openTutorial: true })} className={`flex-1 inline-flex items-center justify-center gap-2 min-w-[148px] px-4 py-2.5 rounded-lg text-sm font-semibold transition-all duration-300 whitespace-nowrap ${activeTab === 'DASHBOARD' ? 'bg-white text-indigo-700 shadow-sm' : 'text-slate-500 hover:bg-white/60 hover:text-slate-700'}`} type="button"><LayoutDashboard className="w-4 h-4" /> <span>Graficos</span></button>
                                    <button id="tour-tab-sheets" onClick={() => handleTabChange('SHEETS', { openTutorial: true })} className={`flex-1 inline-flex items-center justify-center gap-2 min-w-[148px] px-4 py-2.5 rounded-lg text-sm font-semibold transition-all duration-300 whitespace-nowrap ${activeTab === 'SHEETS' ? 'bg-green-600 text-white shadow-sm' : 'text-green-700 hover:bg-white/60'}`} type="button"><TableProperties className="w-4 h-4" /> <span>Planilha</span></button>
                                </div>
                            </div>

                            <div id="tour-header-actions" className="grid grid-cols-2 md:grid-cols-4 xl:grid-cols-[auto_repeat(3,minmax(0,1fr))_auto] items-stretch gap-2">
                                <div className="flex items-center justify-center bg-slate-100 rounded-lg p-1 border border-slate-200 min-h-[46px]">
                                    <button id="tour-undo-btn" onClick={handleUndo} disabled={history.past.length === 0} className="p-1.5 text-slate-600 hover:bg-white hover:shadow-sm rounded-md disabled:opacity-30 disabled:hover:bg-transparent transition-all" type="button"><Undo2 className="w-4 h-4" /></button>
                                    <div className="w-px h-4 bg-slate-300 mx-1" />
                                    <button id="tour-redo-btn" onClick={handleRedo} disabled={history.future.length === 0} className="p-1.5 text-slate-600 hover:bg-white hover:shadow-sm rounded-md disabled:opacity-30 disabled:hover:bg-transparent transition-all" type="button"><Redo2 className="w-4 h-4" /></button>
                                </div>
                                <label id="tour-import-btn" className="cursor-pointer inline-flex w-full items-center justify-center gap-2 px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-sm font-bold transition-all shadow-sm active:scale-95 border border-slate-200 min-h-[46px]">
                                    <UploadCloud className="w-4 h-4" /> <span>Importar</span>
                                    <input type="file" accept=".xlsx, .xls" className="hidden" onChange={handleFileUpload} disabled={loading} />
                                </label>
                                <button id="tour-export-btn" onClick={handleExportExcel} className="inline-flex w-full items-center justify-center gap-2 px-4 py-2.5 bg-slate-800 text-white hover:bg-slate-900 rounded-lg text-sm font-semibold transition-all shadow-md hover:shadow-lg whitespace-nowrap active:scale-95 min-h-[46px]" type="button"><Download className="w-4 h-4" /> <span>Exportar</span></button>
                                <button id="tour-sync-btn" onClick={() => setIsGSheetsModalOpen(true)} className="inline-flex w-full items-center justify-center gap-2 px-4 py-2.5 bg-green-100 text-green-800 hover:bg-green-200 border border-green-300 rounded-lg text-sm font-bold transition-all shadow-sm whitespace-nowrap active:scale-95 min-h-[46px]" type="button"><BarChart2 className="w-4 h-4" /> <span>Atualizar</span></button>
                                {isAuthenticated && (
                                    <div className="inline-flex min-w-[150px] items-center justify-between gap-3 px-3 py-2 rounded-lg border border-slate-200 bg-white shadow-sm min-h-[46px]">
                                        <div className="flex min-w-0 flex-col leading-none text-left">
                                            <span className="text-[11px] font-bold uppercase tracking-wider text-slate-400">Acesso</span>
                                            <span className="text-sm font-semibold text-slate-700 truncate">{getFirstName(currentUsername)}</span>
                                        </div>
                                        <button onClick={handleLogout} className="shrink-0 p-2 bg-white border border-slate-200 hover:bg-slate-100 text-slate-600 rounded-lg transition-all shadow-sm active:scale-95" type="button" aria-label="Sair do sistema">
                                            <LogOut className="w-4 h-4" />
                                        </button>
                                    </div>
                                )}
                            </div>
                        </div>
                    </div>
                </header>

                <main ref={mainScrollRef} className={`flex-1 overflow-y-auto overflow-x-hidden relative w-full bg-slate-50 transition-all duration-400 ease-in-out ${isCinematic ? 'cinematic-effect' : ''}`}>
                    <div className={`${activeTab === 'SHEETS' ? 'max-w-[2200px]' : 'max-w-[1600px]'} mx-auto px-3 sm:px-4 lg:px-6 py-8 w-full relative z-10 min-h-full flex flex-col`}>
                        {data.length === 0 && !isInitialSyncing && (
                            <div className="flex flex-col items-center justify-center flex-1 text-center animate-in fade-in zoom-in-95 duration-500 my-auto">
                                <div className="w-24 h-24 bg-indigo-50 rounded-full flex items-center justify-center mb-6"><Map className="w-12 h-12 text-indigo-400" /></div>
                                <h2 className="text-3xl font-bold text-slate-800 mb-3">Ambiente Zerado</h2>
                                <p className="text-slate-500 max-w-md mx-auto mb-8 leading-relaxed">Inicie sua jornada sincronizando a planilha no topo da tela.</p>
                            </div>
                        )}

                        {activeTab === 'TABELA' && data.length > 0 && (
                            <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden flex flex-col flex-1 animate-in fade-in zoom-in-[0.99] duration-500">
                                <div id="tour-filters" className="p-5 border-b bg-slate-50 border-slate-200 shrink-0">
                                    <div className="mb-4 bg-indigo-50 border border-indigo-200 p-3 rounded-lg flex items-center justify-between shadow-sm animate-in fade-in slide-in-from-top-2 duration-500">
                                        <div className="flex items-center gap-2 text-indigo-800 font-medium text-sm">
                                            <AlertCircle className="w-5 h-5 text-indigo-600" />
                                            A tabela ordena com prioridade vagas abertas em atraso ou urgente.
                                        </div>
                                        <div id="tour-record-count" className="text-xs font-bold text-indigo-600 bg-white px-2 py-1 rounded-md shadow-sm transition-all">Exibindo {filteredData.length} registros</div>
                                    </div>

                                    <div id="tour-filters-controls" className="grid grid-cols-1 md:grid-cols-5 gap-3">
                                        <div id="tour-search" className="relative group"><Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 group-focus-within:text-indigo-500 transition-colors" /><input ref={tableSearchInputRef} type="text" placeholder="Buscar..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="pl-9 pr-4 py-2.5 border border-slate-300 rounded-lg text-sm w-full focus:ring-2 focus:ring-indigo-500 transition-all" /></div>
                                        <select id="tour-status-filter" value={filters.status} onChange={(e) => setFilters((f) => ({ ...f, status: e.target.value }))} className="px-3 py-2.5 border border-slate-300 rounded-lg text-sm focus:ring-2 focus:ring-indigo-500 bg-white font-medium transition-all"><option value="TODOS">Todos os Status</option>{listOptions.status.map((s) => <option key={s} value={s}>{s}</option>)}</select>
                                        <select id="tour-municipio-filter" value={filters.municipio} onChange={(e) => setFilters((f) => ({ ...f, municipio: e.target.value }))} className="px-3 py-2.5 border border-slate-300 rounded-lg text-sm focus:ring-2 focus:ring-indigo-500 bg-white transition-all"><option value="TODOS">Municipios</option>{listOptions.municipios.map((m) => <option key={m} value={m}>{m}</option>)}</select>
                                        <select id="tour-urgencia-filter" value={filters.urgencia} onChange={(e) => setFilters((f) => ({ ...f, urgencia: e.target.value }))} className="px-3 py-2.5 border border-indigo-300 rounded-lg text-sm focus:ring-2 focus:ring-indigo-500 bg-indigo-50 font-bold text-indigo-900 transition-all"><option value="TODOS">Filtro de Prazo</option><option value="URGENTE">Urgente (&lt;= 5d)</option><option value="MEDIA">Em breve (6 a 30d)</option><option value="LONGE">Longo Prazo</option></select>
                                        {showErrorsOnly ? (
                                            <button id="tour-errors-filter" onClick={() => setShowErrorsOnly(false)} className="bg-slate-200 text-slate-700 font-bold text-xs rounded-lg hover:bg-slate-300 transition-all flex items-center justify-center gap-2 active:scale-95" type="button"><X className="w-4 h-4" /> Mostrar Tudo</button>
                                        ) : (
                                            <button id="tour-errors-filter" onClick={() => setShowErrorsOnly(true)} disabled={!metrics?.invalidCount} className={`font-bold text-xs rounded-lg transition-all flex items-center justify-center gap-2 active:scale-95 disabled:cursor-not-allowed disabled:hover:bg-slate-100 ${metrics?.invalidCount > 0 ? 'bg-red-100 text-red-700 hover:bg-red-200 animate-pulse' : 'bg-slate-100 text-slate-400 border border-slate-200'}`} type="button"><AlertCircle className="w-4 h-4" /> {metrics?.invalidCount > 0 ? 'Corrigir Erros' : 'Sem erros agora'}</button>
                                        )}
                                    </div>

                                    <div className="mt-4 flex flex-wrap items-center justify-between gap-3">
                                        <button id="tour-quick-add-btn" onClick={openQuickAddModal} className="inline-flex items-center gap-2 px-5 py-3 rounded-xl bg-emerald-600 text-white text-sm font-bold hover:bg-emerald-700 shadow-sm hover:shadow-md transition-all" type="button"><PlusCircle className="w-5 h-5" /> Novo Cadastro (C)</button>
                                        <div className="text-sm text-slate-700 font-semibold bg-white px-3 py-2 rounded-lg border border-slate-200">Pagina {gridPage} de {gridTotalPages} • Atalho busca: F</div>
                                    </div>
                                </div>

                                <div className="overflow-x-auto relative flex-1 min-h-[400px]">
                                    <table className="w-full text-left text-sm whitespace-nowrap">
                                        <thead id="tour-table-head" className="bg-slate-800 border-b border-slate-700 text-slate-200 font-semibold text-xs uppercase tracking-wider sticky top-0 z-10">
                                            <tr><th id="tour-table-col-status" className="px-6 py-4">Status Rapido</th><th id="tour-table-col-vaga" className="px-6 py-4">Vaga</th><th id="tour-table-col-candidato" className="px-6 py-4">Candidato</th><th id="tour-table-col-prazo" className="px-6 py-4">Prazo</th><th id="tour-table-col-ficha" className="px-6 py-4 text-right">Acoes</th></tr>
                                        </thead>
                                        <tbody id="tour-table" key={tableAnimationKey} className="divide-y divide-slate-200/50">
                                            {gridPagedData.map((row, index) => {
                                                const status = safeGet(row, 'Status');
                                                const candidato = row.Candidato || 'SEM COBERTURA';
                                                const prazoValue = getPrazoValue(row);
                                                const diasParaInicio = getDaysDiff(prazoValue);
                                                return (
                                                    <tr key={row._id} className={`${getRowThermalClass(diasParaInicio, status, candidato, row._isInvalid)} group anim-cascade transition-all duration-300 hover:shadow-sm`} style={{ animationDelay: `${index * 15}ms` }}>
                                                        <>
                                                            <td className="px-6 py-4 relative">
                                                                {row._isInvalid && <AlertCircle className="absolute left-1 top-1/2 -translate-y-1/2 w-4 h-4 text-red-600 animate-pulse" />}
                                                                <select value={status} onChange={(e) => handleInlineEdit(row._id, 'Status', e.target.value)} className={`appearance-none w-32 ml-2 px-2 py-1.5 rounded-lg text-xs font-bold cursor-pointer focus:ring-2 focus:ring-indigo-500 shadow-sm transition-all ${STATUS_COLORS[status] || 'bg-slate-100 text-slate-800 border-slate-200'}`}>
                                                                    <option value="ABERTA">ABERTA</option><option value="FECHADA">FECHADA</option><option value="ENCAMINHADA">ENCAMINHADA</option><option value="CANCELADA">CANCELADA</option><option value="PAUSADA">PAUSADA</option>
                                                                </select>
                                                            </td>
                                                            <td className="px-6 py-4"><div className={`font-bold ${row._isInvalid ? 'text-red-700' : 'text-slate-900'}`}>{row['Nome Subs'] || 'FALTANDO'}</div><div className="text-slate-600 text-xs mt-0.5">{row.CARGO} • {row['NRE / MUNICIPIO']}</div></td>
                                                            <td className="px-6 py-4"><div className={`font-bold ${candidato === 'SEM COBERTURA' ? 'text-red-700' : 'text-slate-800'}`}>{candidato}</div><div className="text-slate-500 text-xs mt-0.5">{row['Contato Candidato'] || 'Sem contato'}</div></td>
                                                            <td className="px-6 py-4"><div className="font-bold text-slate-800">{prazoValue || 'Sem prazo'}</div>
                                                                {diasParaInicio !== null && !['FECHADA', 'ENCAMINHADA', 'CANCELADA'].includes(status) && candidato === 'SEM COBERTURA' && (
                                                                    <div className={`text-xs mt-1 font-bold inline-block px-2 py-0.5 rounded shadow-sm transition-transform group-hover:scale-105 ${diasParaInicio < 0 ? 'bg-red-600 text-white' : diasParaInicio === 0 ? 'bg-red-500 text-white' : diasParaInicio <= 5 ? 'bg-orange-500 text-white' : diasParaInicio <= 15 ? 'bg-yellow-400 text-yellow-900' : diasParaInicio <= 30 ? 'bg-lime-500 text-lime-900' : 'bg-green-500 text-white'}`}>
                                                                        {diasParaInicio < 0 ? `Atrasado ${Math.abs(diasParaInicio)}d` : diasParaInicio === 0 ? 'E Hoje!' : `Faltam ${diasParaInicio}d`}
                                                                    </div>
                                                                )}
                                                            </td>
                                                        </>
                                                        <td className="px-6 py-4 text-right">
                                                            <div className="inline-flex items-center gap-2">
                                                                <button id={index === 0 ? 'tour-record-open-btn' : undefined} onClick={() => { setSelectedRecord(row); setIsEditing(false); setDeleteConfirmOpen(false); }} className="p-2.5 bg-white hover:bg-indigo-50 rounded-lg shadow-sm border border-slate-200 transition-all hover:border-indigo-300 hover:shadow-md active:scale-95" type="button" title="Abrir ficha">
                                                                    <Edit2 className="w-4 h-4 text-slate-600 hover:text-indigo-700" />
                                                                </button>
                                                                <button onClick={() => requestDeleteRecord(row)} className="p-2.5 bg-white hover:bg-red-50 rounded-lg shadow-sm border border-slate-200 transition-all hover:border-red-300 hover:shadow-md active:scale-95" type="button" title="Excluir registro">
                                                                    <Trash2 className="w-4 h-4 text-slate-600 hover:text-red-700" />
                                                                </button>
                                                            </div>
                                                        </td>
                                                    </tr>
                                                );
                                            })}
                                        </tbody>
                                    </table>
                                </div>
                                <div className="px-5 py-3 border-t border-slate-200 bg-slate-50 flex items-center justify-between">
                                    <div className="text-xs font-semibold text-slate-600">Mostrando {gridPagedData.length} itens nesta pagina</div>
                                    <div className="inline-flex items-center gap-2">
                                        <button onClick={() => setGridPage((p) => Math.max(1, p - 1))} disabled={gridPage <= 1} className="px-3 py-1.5 rounded-lg border border-slate-300 bg-white text-sm font-semibold disabled:opacity-40" type="button">Anterior</button>
                                        <button onClick={() => setGridPage((p) => Math.min(gridTotalPages, p + 1))} disabled={gridPage >= gridTotalPages} className="px-3 py-1.5 rounded-lg border border-slate-300 bg-white text-sm font-semibold disabled:opacity-40" type="button">Proxima</button>
                                    </div>
                                </div>
                            </div>
                        )}

                        {activeTab === 'DASHBOARD' && data.length > 0 && metrics && (
                            <div id="tour-dashboard-panel" className="space-y-6 animate-in fade-in zoom-in-[0.98] duration-500">
                                <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
                                    <KpiCard title="Total Cadastrado" value={metrics.total} icon={<FileSpreadsheet className="w-6 h-6 text-indigo-600" />} bg="bg-indigo-50" />
                                    <KpiCard title="Dados Invalidos" value={metrics.invalidCount} icon={<AlertCircle className="w-6 h-6 text-red-600" />} bg={metrics.invalidCount > 0 ? 'bg-red-200 animate-pulse' : 'bg-slate-100'} />
                                    <KpiCard title="Vagas Abertas" value={metrics.abertas} icon={<AlertCircle className="w-6 h-6 text-orange-600" />} bg="bg-orange-50" />
                                    <KpiCard title="Sem Cobertura" value={metrics.semCobertura} icon={<Users className="w-6 h-6 text-amber-600" />} bg="bg-amber-50" />
                                </div>
                                <div className="flex justify-between items-center mt-8 border-b pb-2"><h2 className="font-bold text-slate-800 text-lg">Graficos da operacao</h2><button id="tour-create-chart-btn" onClick={() => setIsChartModalOpen(true)} className="flex items-center gap-2 bg-indigo-50 text-indigo-700 px-3 py-1.5 rounded-lg text-sm font-bold shadow-sm transition-all hover:bg-indigo-100 active:scale-95" type="button"><PlusCircle className="w-4 h-4" /> Criar Grafico</button></div>
                                <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                                    <ChartCard title="Distribuicao de Status" type="pie" data={metrics.statusChartData} groupBy="Status" onChartClick={handleChartClick} />
                                    <ChartCard title="Top 5 Motivos" type="bar" data={metrics.motivoChartData} groupBy="Motivo" onChartClick={handleChartClick} />
                                    {customCharts.map((chart) => <ChartCard key={chart.id} title={chart.title} type={chart.type} data={metrics.generateChartData(chart.groupBy, 10)} groupBy={chart.groupBy} onChartClick={handleChartClick} onDelete={() => removeCustomChart(chart.id)} />)}
                                </div>
                            </div>
                        )}

                        {activeTab === 'SHEETS' && data.length > 0 && (
                            <div id="tour-sheets-table" className="bg-white shadow-xl border border-green-300 rounded-lg flex flex-col flex-1 min-h-[600px] overflow-hidden animate-in fade-in zoom-in-[0.98] duration-500">
                                <div className="bg-green-600 text-white p-3 flex justify-between items-center shadow-sm z-10 shrink-0">
                                    <div className="flex items-center gap-2"><TableProperties className="w-5 h-5 text-green-100" /><h2 className="font-bold">Planilha editável</h2></div>
                                    <div className="text-sm font-semibold">Exibindo {sheetFilteredData.length} de {filteredData.length} registros</div>
                                </div>

                                {/* Barra de Filtros da Planilha */}
                                <div className="bg-green-50 border-b border-green-200 p-3 z-10 shrink-0">
                                    <div id="tour-sheets-dynamic-filters" className="grid grid-cols-1 md:grid-cols-4 gap-2 items-center mb-2">
                                        <div className="relative">
                                            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                                            <input
                                                ref={sheetSearchInputRef}
                                                type="text"
                                                placeholder="Buscar em todas as colunas..."
                                                value={sheetSearchTerm}
                                                onChange={(e) => setSheetSearchTerm(e.target.value)}
                                                className="w-full pl-9 pr-3 py-2 border border-green-300 rounded-lg text-sm focus:ring-2 focus:ring-green-500 focus:border-transparent bg-white"
                                            />
                                        </div>
                                        <select
                                            value={sheetFilterColumn}
                                            onChange={(e) => setSheetFilterColumn(e.target.value)}
                                            className="w-full px-3 py-2 border border-green-300 rounded-lg text-sm bg-white focus:ring-2 focus:ring-green-500"
                                        >
                                            <option value="TODOS">Filtrar por coluna...</option>
                                            {sheetsColumns.map((column) => <option key={`sheet-filter-col-${column}`} value={column}>{column}</option>)}
                                        </select>
                                        <input
                                            type="text"
                                            value={sheetFilterTerm}
                                            onChange={(e) => setSheetFilterTerm(e.target.value)}
                                            placeholder="Termo na coluna"
                                            className="w-full px-3 py-2 border border-green-300 rounded-lg text-sm bg-white focus:ring-2 focus:ring-green-500"
                                        />
                                        <button
                                            onClick={() => { setSheetSearchTerm(''); setSheetFilterColumn('TODOS'); setSheetFilterTerm(''); }}
                                            disabled={!sheetSearchTerm && !(sheetFilterColumn !== 'TODOS' && sheetFilterTerm.trim())}
                                            className="px-3 py-2 bg-red-100 text-red-700 rounded-lg text-sm font-medium hover:bg-red-200 transition-all disabled:opacity-40 disabled:hover:bg-red-100"
                                            type="button"
                                        >
                                            Limpar
                                        </button>
                                    </div>

                                    <p className="text-xs text-slate-600">Busca rapida: use a barra global ou escolha uma coluna + termo.</p>

                                    <div className="mt-3 flex items-center justify-end gap-2">
                                        <div ref={sheetColumnsPanelRef} className="relative">
                                            <button id="tour-sheets-columns-panel-btn" onClick={() => setIsSheetColumnsPanelOpen((open) => !open)} className="inline-flex items-center gap-2 px-4 py-2.5 rounded-xl bg-white text-slate-700 text-sm font-bold hover:bg-slate-100 border border-slate-300 shadow-sm transition-all" type="button">
                                                <Sliders className="w-4 h-4" /> Colunas
                                            </button>

                                            {isSheetColumnsPanelOpen && (
                                                <div className="absolute right-0 mt-2 w-[320px] max-h-[360px] overflow-hidden bg-white border border-slate-200 rounded-2xl shadow-xl z-20">
                                                    <div className="px-4 py-3 border-b bg-slate-50 flex items-center justify-between">
                                                        <p className="text-xs font-bold uppercase tracking-wider text-slate-500">Exibir colunas</p>
                                                        <button onClick={resetSheetColumnsLayout} className="text-xs font-bold text-indigo-600 hover:text-indigo-800" type="button">Resetar layout</button>
                                                    </div>
                                                    <div className="p-3 space-y-1 max-h-[280px] overflow-auto">
                                                        {sheetsColumns.map((column, index) => {
                                                            const checked = !hiddenSheetColumns.includes(column);
                                                            return (
                                                                <label id={index === 0 ? 'tour-sheets-columns-first-toggle' : undefined} key={`sheet-col-toggle-${column}`} className="flex items-center gap-2 px-2 py-1.5 rounded-lg hover:bg-slate-50 cursor-pointer">
                                                                    <input
                                                                        type="checkbox"
                                                                        checked={checked}
                                                                        onChange={() => toggleSheetColumnVisibility(column)}
                                                                        disabled={checked && visibleSheetColumns.length <= 1}
                                                                        className="rounded text-green-600 focus:ring-green-500"
                                                                    />
                                                                    <span className="text-sm text-slate-700 truncate">{column}</span>
                                                                </label>
                                                            );
                                                        })}
                                                    </div>
                                                </div>
                                            )}
                                        </div>

                                        <button
                                            onClick={() => {
                                                setActiveTab('SHEETS');
                                                setTutorialSection('NEW_FEATURES_SHEETS_V2');
                                                setShowTutorial(true);
                                            }}
                                            className="inline-flex items-center gap-2 px-4 py-2.5 rounded-xl bg-indigo-600 text-white text-sm font-bold hover:bg-indigo-700 border border-indigo-600 shadow-sm transition-all"
                                            type="button"
                                        >
                                            <ChevronRight className="w-4 h-4" /> Tutorial Planilha
                                        </button>

                                        <button onClick={openQuickAddModal} className="inline-flex items-center gap-2 px-5 py-3 rounded-xl bg-emerald-600 text-white text-sm font-bold hover:bg-emerald-700 shadow-sm hover:shadow-md transition-all" type="button"><PlusCircle className="w-5 h-5" /> Novo Cadastro (C)</button>
                                    </div>
                                </div>

                                <div id="tour-sheets-horizontal-scroll" className="overflow-auto flex-1 bg-slate-100 p-2">
                                    <table className="w-max min-w-full table-fixed text-left text-xs border-collapse bg-white shadow-sm ring-1 ring-slate-200">
                                        <colgroup>
                                            {visibleSheetColumns.map((column) => (
                                                <col key={`sheet-col-${column}`} style={{ width: `${sheetColumnWidths[column] || 100}px` }} />
                                            ))}
                                            <col style={{ width: '72px' }} />
                                        </colgroup>
                                        <thead id="tour-sheets-head" className="bg-slate-100 border-b-2 border-slate-300 text-slate-700 font-bold text-xs sticky top-0 z-10 shadow-sm">
                                            <tr>
                                                {visibleSheetColumns.map((column, index) => {
                                                    const isLast = index === visibleSheetColumns.length - 1;
                                                    const isCandidateBlock = ['Candidato', 'Contato Candidato'].includes(column);
                                                    return (
                                                        <th
                                                            key={column}
                                                            title={column}
                                                            className={`px-2 py-2 truncate relative select-none ${isLast ? '' : 'border-r'} ${isCandidateBlock ? 'bg-green-50' : ''}`}
                                                        >
                                                            {column}
                                                            <div
                                                                onMouseDown={(event) => startSheetColumnResize(column, event)}
                                                                className="absolute right-0 top-0 h-full w-2 cursor-col-resize hover:bg-green-300/50"
                                                                title="Arraste para ajustar largura"
                                                            />
                                                        </th>
                                                    );
                                                })}
                                                <th id="tour-sheets-actions-col" className="px-2 py-2 w-[72px] text-right">Acoes</th>
                                            </tr>
                                        </thead>
                                        <tbody key={`sheets-${tableAnimationKey}`} className="divide-y divide-slate-200 font-medium">
                                            {sheetFilteredData.length === 0 ? (
                                                <tr>
                                                    <td colSpan={visibleSheetColumns.length + 1} className="px-4 py-8 text-center text-slate-500 font-medium">
                                                        Nenhum registro encontrado com os filtros aplicados
                                                    </td>
                                                </tr>
                                            ) : (
                                                sheetPagedData.map((row) => (
                                                    <tr key={row._id} className={`hover:bg-blue-50/50 transition-colors ${row._isInvalid ? 'bg-red-50' : ''}`}>
                                                        {visibleSheetColumns.map((column, index) => {
                                                            const isLast = index === visibleSheetColumns.length - 1;
                                                            const isCandidateBlock = ['Candidato', 'Contato Candidato'].includes(column);
                                                            const isStatus = normalizeCredentialText(column) === 'status';
                                                            const isRequiredName = normalizeCredentialText(column) === normalizeCredentialText('Nome Subs');
                                                            const rawValue = row[column];
                                                            const value = rawValue === undefined || rawValue === null ? '' : String(rawValue);

                                                            return (
                                                                <td key={`${row._id}-${column}`} className={`${isLast ? '' : 'border-r'} p-0 min-w-0 ${isCandidateBlock ? 'bg-green-50/30' : ''} ${row._isInvalid && isRequiredName && !value.trim() ? 'ring-2 ring-inset ring-red-500 bg-red-50' : ''}`}>
                                                                    {isStatus ? (
                                                                        <select value={value} onChange={(e) => handleInlineEdit(row._id, column, e.target.value)} className="w-full h-full px-2 py-1.5 text-xs appearance-none cursor-pointer focus:bg-blue-100 bg-transparent transition-colors" title={value}>
                                                                            <option value="ABERTA">ABERTA</option><option value="FECHADA">FECHADA</option><option value="ENCAMINHADA">ENCAMINHADA</option><option value="CANCELADA">CANCELADA</option><option value="PAUSADA">PAUSADA</option>
                                                                        </select>
                                                                    ) : (
                                                                        <input
                                                                            type="text"
                                                                            value={value}
                                                                            onChange={(e) => handleInlineEdit(row._id, column, e.target.value)}
                                                                            title={value}
                                                                            className={`w-full h-full px-2 py-1.5 text-xs bg-transparent transition-colors ${isCandidateBlock ? 'focus:bg-green-100' : 'focus:bg-blue-100'} ${normalizeCredentialText(column) === normalizeCredentialText('OBS:') ? 'text-[11px]' : ''} ${normalizeCredentialText(column) === normalizeCredentialText('Candidato') ? 'font-bold text-green-900 placeholder:text-green-300' : ''} ${row._isInvalid && isRequiredName ? 'placeholder:text-red-300' : ''}`}
                                                                            placeholder={row._isInvalid && isRequiredName ? 'Obrigatorio' : ''}
                                                                        />
                                                                    )}
                                                                </td>
                                                            );
                                                        })}
                                                        <td className="px-2 py-1 text-right border-l border-slate-200 bg-white align-middle">
                                                            <button
                                                                onClick={() => requestDeleteRecord(row)}
                                                                className="p-2 bg-white hover:bg-red-50 rounded-lg shadow-sm border border-slate-200 transition-all hover:border-red-300 hover:shadow-md active:scale-95"
                                                                type="button"
                                                                title="Excluir linha"
                                                            >
                                                                <Trash2 className="w-4 h-4 text-slate-600 hover:text-red-700" />
                                                            </button>
                                                        </td>
                                                    </tr>
                                                ))
                                            )}
                                        </tbody>
                                    </table>
                                </div>
                                <div className="px-4 py-3 border-t border-slate-200 bg-slate-50 flex items-center justify-between">
                                    <div className="text-xs font-semibold text-slate-600">Pagina {sheetPage} de {sheetTotalPages} • {sheetPagedData.length} linhas visiveis</div>
                                    <div className="inline-flex items-center gap-2">
                                        <button onClick={() => setSheetPage((p) => Math.max(1, p - 1))} disabled={sheetPage <= 1} className="px-3 py-1.5 rounded-lg border border-slate-300 bg-white text-sm font-semibold disabled:opacity-40" type="button">Anterior</button>
                                        <button onClick={() => setSheetPage((p) => Math.min(sheetTotalPages, p + 1))} disabled={sheetPage >= sheetTotalPages} className="px-3 py-1.5 rounded-lg border border-slate-300 bg-white text-sm font-semibold disabled:opacity-40" type="button">Proxima</button>
                                    </div>
                                </div>
                            </div>
                        )}
                    </div>
                </main>
            </div>

            {showPatchNotes && !showWelcome && isAuthenticated && (
                <PatchNotesModal
                    onClose={handlePatchNotesClose}
                    onStartTableTutorial={handlePatchNotesStartTableTutorial}
                    onStartSheetsTutorial={handlePatchNotesStartSheetsTutorial}
                />
            )}

            {deleteModalRecord && (
                <div className="fixed inset-0 z-[210] bg-slate-900/65 backdrop-blur-sm flex items-center justify-center p-4 animate-in fade-in duration-200">
                    <div className="w-full max-w-md bg-white rounded-3xl border border-slate-200 shadow-2xl overflow-hidden animate-in zoom-in-95 duration-200">
                        <div className="px-6 py-5 border-b bg-gradient-to-r from-red-50 to-rose-50 flex items-start gap-3">
                            <div className="p-2.5 rounded-2xl bg-red-100">
                                <Trash2 className="w-5 h-5 text-red-600" />
                            </div>
                            <div className="min-w-0">
                                <h3 className="text-lg font-black text-slate-800">Confirmar exclusao</h3>
                                <p className="text-sm text-slate-600 mt-0.5">Essa acao remove o registro da tabela e da planilha.</p>
                            </div>
                        </div>

                        <div className="px-6 py-5 space-y-3 bg-white">
                            <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4 space-y-2">
                                <p className="text-[11px] uppercase tracking-wider font-bold text-slate-500">Registro selecionado</p>
                                <p className="text-sm font-bold text-slate-800 break-words">{String(getRowValue(deleteModalRecord, ['Nome Subs']) || 'Nome nao informado')}</p>
                                <div className="flex items-center gap-2">
                                    <StatusBadge status={String(getRowValue(deleteModalRecord, ['Status']) || '')} />
                                    <span className="text-xs text-slate-500">{String(getRowValue(deleteModalRecord, ['CARGO']) || 'Sem cargo')}</span>
                                </div>
                            </div>

                            <p className="text-sm text-slate-600">Deseja mesmo excluir este registro?</p>
                        </div>

                        <div className="px-6 py-4 border-t bg-slate-50 flex items-center justify-end gap-3">
                            <button onClick={handleDeleteModalCancel} className="px-5 py-2.5 rounded-xl text-slate-600 font-bold hover:bg-slate-200 transition-colors" type="button">Cancelar</button>
                            <button onClick={handleDeleteModalConfirm} className="px-5 py-2.5 rounded-xl bg-red-600 text-white font-bold hover:bg-red-700 shadow-md active:scale-95 transition-all inline-flex items-center gap-2" type="button">
                                <Trash2 className="w-4 h-4" />
                                Excluir agora
                            </button>
                        </div>
                    </div>
                </div>
            )}

            {showTutorial && !showWelcome && tutorialSteps.length > 0 && (
                <WalkthroughTour
                    key={tutorialSection}
                    section={tutorialSection}
                    steps={tutorialSteps}
                    onStepChange={setTutorialActiveStep}
                    onComplete={(dontShow) => handleTutorialComplete(tutorialSection, dontShow)}
                />
            )}

            {isGSheetsModalOpen && (
                <div className="fixed inset-0 z-50 bg-slate-900/60 flex items-center justify-center p-4 backdrop-blur-sm animate-in fade-in duration-200">
                    <div className="bg-white rounded-3xl p-8 max-w-md w-full shadow-2xl relative animate-in zoom-in-95 duration-300">
                        <button onClick={() => setIsGSheetsModalOpen(false)} className="absolute top-4 right-4 text-slate-400 hover:text-slate-800 transition-colors" type="button"><X className="w-5 h-5" /></button>
                        <h3 className="text-xl font-bold text-slate-800 mb-4 flex items-center gap-2"><TableProperties className="w-6 h-6 text-green-600" /> Sincronizar Google Sheets</h3>
                        <code className="block w-full p-3 bg-slate-100 rounded-lg text-xs overflow-x-auto mb-6 border font-mono">{GOOGLE_SHEETS_URL}</code>
                        <button onClick={handleGoogleSheetsSync} className="w-full bg-green-600 hover:bg-green-700 text-white font-bold py-3 rounded-xl flex items-center justify-center gap-2 shadow-md transition-all active:scale-95" type="button">
                            <BarChart2 className={`w-5 h-5 ${loading ? 'animate-spin' : ''}`} /> Sincronizar Agora
                        </button>
                    </div>
                </div>
            )}

            {isChartModalOpen && (
                <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4 animate-in fade-in duration-200">
                    <div id="tour-chart-modal" className="bg-white rounded-3xl shadow-2xl w-full max-w-md p-8 animate-in zoom-in-95 duration-300">
                        <h3 className="text-xl font-bold text-slate-800 mb-6 flex items-center gap-2 border-b pb-3"><BarChart2 className="w-6 h-6 text-indigo-600" /> Montador de Graficos</h3>
                        <div className="space-y-5">
                            <div><label className="text-xs font-bold text-slate-500 uppercase">Titulo Visual</label><input type="text" value={newChartData.title} onChange={(e) => setNewChartData({ ...newChartData, title: e.target.value })} className="w-full mt-1.5 px-4 py-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-indigo-500 transition-all shadow-sm" /></div>
                            <div><label className="text-xs font-bold text-slate-500 uppercase">Ancorar Dados em</label><select value={newChartData.groupBy} onChange={(e) => setNewChartData({ ...newChartData, groupBy: e.target.value })} className="w-full mt-1.5 px-4 py-2.5 border border-slate-300 rounded-xl bg-white font-medium text-slate-700 transition-all shadow-sm">{availableColumns.map((col) => <option key={col} value={col}>{col}</option>)}</select></div>
                            <div>
                                <label className="text-xs font-bold text-slate-500 uppercase">Estilo Grafico</label>
                                <div className="flex gap-3 mt-1.5">
                                    <button onClick={() => setNewChartData({ ...newChartData, type: 'bar' })} className={`flex-1 py-3 flex justify-center gap-2 rounded-xl border font-bold text-sm transition-all ${newChartData.type === 'bar' ? 'bg-indigo-50 border-indigo-500 text-indigo-700 shadow-inner' : 'bg-white hover:bg-slate-50'}`} type="button"><BarChart2 className="w-4 h-4" /> Barras</button>
                                    <button onClick={() => setNewChartData({ ...newChartData, type: 'pie' })} className={`flex-1 py-3 flex justify-center gap-2 rounded-xl border font-bold text-sm transition-all ${newChartData.type === 'pie' ? 'bg-indigo-50 border-indigo-500 text-indigo-700 shadow-inner' : 'bg-white hover:bg-slate-50'}`} type="button"><PieChartIcon className="w-4 h-4" /> Pizza</button>
                                </div>
                            </div>
                        </div>
                        <div className="flex justify-end gap-3 mt-8"><button onClick={() => setIsChartModalOpen(false)} className="px-5 py-2 text-slate-500 font-bold hover:bg-slate-100 rounded-xl transition-colors" type="button">Cancelar</button><button onClick={addCustomChart} className="px-6 py-2 bg-indigo-600 text-white font-bold rounded-xl hover:bg-indigo-700 shadow-md active:scale-95 transition-all" type="button">Salvar Dashboard</button></div>
                    </div>
                </div>
            )}

            {isQuickAddModalOpen && (
                <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[80] flex items-center justify-center p-4 animate-in fade-in duration-200">
                    <div id="tour-quick-add-modal" className="bg-white rounded-3xl shadow-2xl w-full max-w-4xl max-h-[90vh] flex flex-col overflow-hidden animate-in zoom-in-95 duration-300">
                        <div className="px-8 py-5 border-b flex items-center justify-between bg-white">
                            <h3 className="text-xl font-bold text-slate-800 flex gap-2 items-center"><PlusCircle className="w-6 h-6 text-emerald-600" /> Novo Cadastro Rápido</h3>
                            <button onClick={() => setIsQuickAddModalOpen(false)} className="p-2 text-slate-400 hover:bg-slate-100 rounded-xl transition-colors" type="button"><X className="w-5 h-5" /></button>
                        </div>

                        <div className="p-6 md:p-8 overflow-y-auto flex-1 bg-slate-50">
                            <p className="text-sm text-slate-600 mb-4">Cadastro completo em tela única. Campos com <span className="text-red-600 font-bold">*</span> são obrigatórios.</p>

                            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4">
                                {detailedColumns.map((column) => {
                                    const isStatus = normalizeCredentialText(column) === 'status';
                                    const value = quickAddData[column] === undefined || quickAddData[column] === null ? '' : String(quickAddData[column]);
                                    const isRequired = REQUIRED_FIELDS.some((field) => normalizeCredentialText(field) === normalizeCredentialText(column));

                                    return (
                                        <div key={`new-all-${column}`} className="space-y-1.5">
                                            <label className="text-xs font-bold text-slate-500 uppercase">{column} {isRequired ? <span className="text-red-600">*</span> : null}</label>
                                            {isStatus ? (
                                                <div className="flex flex-wrap gap-2 col-span-full sm:col-span-1">
                                                    {['ABERTA', 'FECHADA', 'ENCAMINHADA', 'CANCELADA', 'PAUSADA'].map((s) => {
                                                        const Icon = STATUS_ICON_MAP[s];
                                                        const isActive = (value || 'ABERTA') === s;
                                                        return (
                                                            <button key={s} type="button" onClick={() => setQuickAddData((prev) => ({ ...prev, [column]: s }))} className={`flex items-center gap-1.5 px-3 py-2 rounded-xl border text-xs font-bold transition-all active:scale-95 ${isActive ? (STATUS_COLORS[s] || 'bg-slate-100 text-slate-700 border-slate-300') : 'bg-slate-50 text-slate-500 border-slate-200 hover:bg-slate-100'}`}>
                                                                {Icon && <Icon className="w-3.5 h-3.5" />}
                                                                {s}
                                                            </button>
                                                        );
                                                    })}
                                                </div>
                                            ) : (
                                                <input type="text" value={value} onChange={(e) => setQuickAddData((prev) => ({ ...prev, [column]: e.target.value }))} className={`w-full px-3 py-2.5 border rounded-xl focus:ring-2 focus:ring-emerald-500 bg-white ${isRequired && !value.trim() ? 'border-red-300' : 'border-slate-300'}`} />
                                            )}
                                        </div>
                                    );
                                })}
                            </div>
                        </div>

                        <div className="px-8 py-4 border-t bg-white flex justify-end gap-3">
                            <button onClick={() => setIsQuickAddModalOpen(false)} className="px-5 py-2 text-slate-500 font-bold hover:bg-slate-100 rounded-xl transition-colors" type="button">Cancelar</button>
                            <button onClick={handleQuickAddSave} className="px-6 py-2 bg-emerald-600 text-white font-bold rounded-xl hover:bg-emerald-700 shadow-md active:scale-95 transition-all" type="button">Salvar Cadastro</button>
                        </div>
                    </div>
                </div>
            )}

            {selectedRecord && (
                <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[90] animate-in fade-in duration-200">
                    <div id="tour-record-modal" className="bg-white rounded-3xl shadow-2xl w-full max-w-3xl max-h-[90vh] flex flex-col overflow-hidden animate-in zoom-in-95 duration-300">

                        {/* Tira de cor por status */}
                        <div className={`h-1.5 shrink-0 ${STATUS_ACCENT[String(getRowValue(selectedRecord, ['Status', 'STATUS']) || '').toUpperCase()] || 'bg-indigo-500'}`} />

                        {/* Cabeçalho */}
                        <div className="px-7 py-5 border-b bg-white flex items-center gap-4">
                            <div className="bg-indigo-50 p-2.5 rounded-2xl shrink-0">
                                <Briefcase className="w-5 h-5 text-indigo-600" />
                            </div>
                            <div className="flex-1 min-w-0">
                                <p className="text-[10px] font-bold uppercase tracking-widest text-slate-400 mb-0.5">{isEditing ? 'Editando ficha' : 'Ficha Completa'}</p>
                                <h3 className="text-lg font-black text-slate-800 truncate">{String(getRowValue(selectedRecord, ['Nome Subs', 'NOME SUBS']) || '—')}</h3>
                            </div>
                            <div className="flex items-center gap-2 shrink-0 flex-wrap justify-end">
                                {!isEditing ? (
                                    <>
                                        <button id="tour-record-edit-btn" onClick={handleEditClick} className="flex items-center gap-1.5 px-4 py-2 bg-indigo-50 text-indigo-700 font-bold rounded-xl text-sm hover:bg-indigo-100 transition-colors" type="button">
                                            <Edit2 className="w-4 h-4" /> Editar
                                        </button>
                                        {deleteConfirmOpen ? (
                                            <div className="flex items-center gap-2 bg-red-50 border border-red-200 rounded-xl px-3 py-2">
                                                <span className="text-xs font-bold text-red-700">Excluir registro?</span>
                                                <button onClick={handleDeleteRecord} className="px-2.5 py-1 bg-red-600 text-white text-xs font-bold rounded-lg hover:bg-red-700 transition-colors active:scale-95" type="button">Sim</button>
                                                <button onClick={() => setDeleteConfirmOpen(false)} className="px-2.5 py-1 bg-white border border-slate-300 text-slate-600 text-xs font-bold rounded-lg hover:bg-slate-50 transition-colors" type="button">Não</button>
                                            </div>
                                        ) : (
                                            <button onClick={() => setDeleteConfirmOpen(true)} className="flex items-center gap-1.5 px-4 py-2 bg-red-50 text-red-700 font-bold rounded-xl text-sm hover:bg-red-100 transition-colors" type="button">
                                                <Trash2 className="w-4 h-4" /> Excluir
                                            </button>
                                        )}
                                    </>
                                ) : (
                                    <button id="tour-record-save-btn" onClick={handleSaveEdit} className="flex items-center gap-1.5 px-5 py-2 bg-green-600 text-white font-bold rounded-xl text-sm hover:bg-green-700 shadow-md transition-all active:scale-95" type="button">
                                        <Save className="w-4 h-4" /> Salvar
                                    </button>
                                )}
                                <button onClick={() => { setSelectedRecord(null); setDeleteConfirmOpen(false); }} className="p-2 text-slate-400 hover:bg-slate-100 rounded-xl transition-colors ml-1" type="button">
                                    <X className="w-5 h-5" />
                                </button>
                            </div>
                        </div>

                        {/* Corpo */}
                        <div className="p-6 md:p-8 overflow-y-auto flex-1 bg-slate-50">
                            {!isEditing ? (
                                <div className="space-y-4">
                                    {/* Card de status em destaque */}
                                    <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-4 flex items-center gap-5">
                                        <div>
                                            <p className="text-[10px] font-bold uppercase tracking-widest text-slate-400 mb-1.5">Status Atual</p>
                                            <StatusBadge status={String(getRowValue(selectedRecord, ['Status', 'STATUS']) || '')} size="lg" />
                                        </div>
                                        {String(getRowValue(selectedRecord, ['CARGO', 'Cargo', 'cargo']) || '') && (
                                            <div className="border-l border-slate-200 pl-5">
                                                <p className="text-[10px] font-bold uppercase tracking-widest text-slate-400 mb-1">Cargo</p>
                                                <p className="text-sm font-bold text-slate-700">{String(getRowValue(selectedRecord, ['CARGO', 'Cargo', 'cargo']))}</p>
                                            </div>
                                        )}
                                        {String(getRowValue(selectedRecord, ['NRE / MUNICIPIO', 'NRE/MUNICIPIO', 'NRE', 'Municipio']) || '') && (
                                            <div className="border-l border-slate-200 pl-5">
                                                <p className="text-[10px] font-bold uppercase tracking-widest text-slate-400 mb-1">Município</p>
                                                <p className="text-sm font-bold text-slate-700">{String(getRowValue(selectedRecord, ['NRE / MUNICIPIO', 'NRE/MUNICIPIO', 'NRE', 'Municipio']))}</p>
                                            </div>
                                        )}
                                    </div>

                                    {/* Grid de todos os campos */}
                                    <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-5">
                                        <h4 className="text-[10px] font-bold uppercase tracking-widest text-slate-400 border-b border-slate-100 pb-2 mb-4">Todos os campos</h4>
                                        <div className="grid grid-cols-1 md:grid-cols-2 gap-2.5">
                                            {detailedColumns.map((column) => {
                                                const fieldValue = String(selectedRecord?.[column] || '');
                                                const isEmpty = !fieldValue.trim();
                                                const isStatusField = normalizeCredentialText(column) === 'status';
                                                return (
                                                    <div key={`view-${column}`} className="p-3 rounded-xl bg-slate-50 border border-slate-100 hover:border-slate-200 transition-colors">
                                                        <p className="text-[10px] font-bold uppercase tracking-widest text-slate-400 mb-1.5">{column}</p>
                                                        {isStatusField ? (
                                                            <StatusBadge status={fieldValue} />
                                                        ) : (
                                                            <p className={`text-sm font-semibold break-words leading-snug ${isEmpty ? 'text-slate-400 italic' : 'text-slate-800'}`}>
                                                                {isEmpty ? 'Não informado' : fieldValue}
                                                            </p>
                                                        )}
                                                    </div>
                                                );
                                            })}
                                        </div>
                                    </div>
                                </div>
                            ) : (
                                <div id="tour-record-edit-fields" className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6">
                                    <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
                                        {detailedColumns.map((column) => {
                                            const isStatus = normalizeCredentialText(column) === 'status';
                                            const value = editFormData[column] === undefined || editFormData[column] === null ? '' : String(editFormData[column]);
                                            return (
                                                <div key={`edit-${column}`} className="space-y-1.5">
                                                    <label className="text-[10px] font-bold uppercase tracking-widest text-slate-400">{column}</label>
                                                    {isStatus ? (
                                                        <div className="flex flex-wrap gap-2">
                                                            {['ABERTA', 'FECHADA', 'ENCAMINHADA', 'CANCELADA', 'PAUSADA'].map((s) => {
                                                                const Icon = STATUS_ICON_MAP[s];
                                                                const isActive = value === s;
                                                                return (
                                                                    <button key={s} type="button" onClick={() => setEditFormData({ ...editFormData, [column]: s })} className={`flex items-center gap-1.5 px-3 py-1.5 rounded-xl border text-xs font-bold transition-all active:scale-95 ${isActive ? (STATUS_COLORS[s] || 'bg-slate-100 text-slate-700 border-slate-300') : 'bg-slate-50 text-slate-500 border-slate-200 hover:bg-slate-100'}`}>
                                                                        {Icon && <Icon className="w-3.5 h-3.5" />}
                                                                        {s}
                                                                    </button>
                                                                );
                                                            })}
                                                        </div>
                                                    ) : (
                                                        <input type="text" value={value} onChange={(e) => setEditFormData({ ...editFormData, [column]: e.target.value })} className="w-full px-4 py-2.5 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 bg-slate-50 text-sm" />
                                                    )}
                                                </div>
                                            );
                                        })}
                                    </div>
                                </div>
                            )}
                        </div>
                    </div>
                </div>
            )}
        </React.Fragment>
    );
}

function KpiCard({ title, value, icon, bg }) {
    return (
        <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 flex flex-col justify-between hover:-translate-y-1 hover:shadow-md transition-all duration-300">
            <div className="flex justify-between mb-4"><div className={`p-3 rounded-2xl ${bg}`}>{icon}</div></div>
            <div><h3 className="text-4xl font-bold text-slate-800 tracking-tight">{value}</h3><p className="text-slate-500 text-sm font-bold uppercase mt-1 tracking-wider">{title}</p></div>
        </div>
    );
}

function ChartCard({ title, type, data, onDelete, onChartClick, groupBy }) {
    return (
        <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 relative group flex flex-col w-full h-full min-h-[320px] hover:shadow-md hover:border-indigo-200 transition-all duration-300">
            {onDelete && <button onClick={onDelete} className="absolute top-4 right-4 p-2 text-slate-300 hover:text-red-500 hover:bg-red-50 rounded-xl opacity-0 group-hover:opacity-100 transition-all active:scale-95" type="button"><Trash2 className="w-4 h-4" /></button>}
            <h2 className="text-sm font-bold text-slate-500 uppercase tracking-wider mb-6 text-center truncate pr-8">{title}</h2>
            {data.length === 0 ? (<div className="flex-1 flex items-center justify-center text-slate-400 text-sm">Sem dados</div>) : (
                <div style={{ width: '100%', height: 250 }} className="animate-in fade-in zoom-in-[0.95] duration-700">
                    <ResponsiveContainer width="100%" height="100%">
                        {type === 'pie' ? (
                            <PieChart>
                                <Pie data={data} cx="50%" cy="50%" innerRadius={55} outerRadius={90} paddingAngle={4} dataKey="value" onClick={(entry) => onChartClick && onChartClick(groupBy, entry.name)} style={{ cursor: onChartClick ? 'pointer' : 'default', outline: 'none' }} className="transition-all duration-300">
                                    {data.map((entry, index) => <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} className="hover:opacity-80 transition-opacity drop-shadow-sm" />)}
                                </Pie>
                                <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.1)' }} cursor={{ fill: 'transparent' }} />
                            </PieChart>
                        ) : (
                            <BarChart data={data} layout="vertical" margin={{ left: -15, right: 10, top: 0, bottom: 0 }}>
                                <XAxis type="number" hide />
                                <YAxis dataKey="name" type="category" width={110} tick={{ fontSize: 11, fill: '#64748b', fontWeight: 600 }} axisLine={false} tickLine={false} />
                                <Tooltip cursor={{ fill: 'transparent' }} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.1)' }} />
                                <Bar dataKey="value" fill="#3b82f6" radius={[0, 6, 6, 0]} onClick={(entry) => onChartClick && onChartClick(groupBy, entry.name)} style={{ cursor: onChartClick ? 'pointer' : 'default' }} className="hover:opacity-80 transition-opacity">
                                    {data.map((entry, index) => <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />)}
                                </Bar>
                            </BarChart>
                        )}
                    </ResponsiveContainer>
                </div>
            )}
        </div>
    );
}
