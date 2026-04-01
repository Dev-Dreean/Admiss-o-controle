import React, { useState, useMemo, useEffect, useRef } from 'react';
import {
    UploadCloud, FileSpreadsheet, AlertCircle, Users, Search,
    X, ChevronRight, Briefcase, Download, Edit2, Save,
    LayoutDashboard, ListTodo, TableProperties, PlusCircle,
    BarChart2, PieChart as PieChartIcon, Trash2,
    Undo2, Redo2, Check, ChevronLeft, Map
} from 'lucide-react';
import {
    PieChart, Pie, Cell, BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer
} from 'recharts';

const COLORS = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899', '#14b8a6', '#f43f5e', '#84cc16'];
const STATUS_COLORS = {
    ABERTA: 'bg-red-100 text-red-800 border-red-300 font-bold',
    FECHADA: 'bg-green-100 text-green-800 border-green-300',
    ENCAMINHADA: 'bg-blue-100 text-blue-800 border-blue-300',
    CANCELADA: 'bg-gray-100 text-gray-800 border-gray-300',
    PAUSADA: 'bg-yellow-100 text-yellow-800 border-yellow-300',
};

const GOOGLE_SHEETS_URL = 'https://docs.google.com/spreadsheets/d/1hmLkIX2B4rh6NDtJUXOhtjdXhddozqPs9uMTzaTeBsk/edit?usp=sharing';
const GOOGLE_SHEETS_CSV_EXPORT = 'https://docs.google.com/spreadsheets/d/1hmLkIX2B4rh6NDtJUXOhtjdXhddozqPs9uMTzaTeBsk/export?format=csv';

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

const DEFAULT_TUTORIAL_PROGRESS = Object.freeze({
    TABELA: false,
    DASHBOARD: false,
    SHEETS: false,
});

const normalizeTutorialProgress = (value) => ({
    TABELA: Boolean(value?.TABELA),
    DASHBOARD: Boolean(value?.DASHBOARD),
    SHEETS: Boolean(value?.SHEETS),
});

const TUTORIAL_STEPS = {
    TABELA: [
        {
            target: 'tour-tabs',
            title: 'Menu principal',
            desc: 'Use estes botoes para trocar entre a tabela, os graficos e a planilha.',
            icon: <LayoutDashboard className="w-8 h-8 text-indigo-500" />,
        },
        {
            target: 'tour-search',
            title: 'Busca',
            desc: 'Digite qualquer nome, cargo ou municipio para achar mais rapido o que voce precisa.',
            icon: <Search className="w-8 h-8 text-orange-500" />,
        },
        {
            target: 'tour-status-filter',
            title: 'Filtro por status',
            desc: 'Aqui voce mostra somente vagas abertas, fechadas, pausadas ou qualquer outro status.',
            icon: <ListTodo className="w-8 h-8 text-indigo-500" />,
        },
        {
            target: 'tour-municipio-filter',
            title: 'Filtro por municipio',
            desc: 'Use este campo para olhar apenas a cidade ou regional que voce quer acompanhar.',
            icon: <Map className="w-8 h-8 text-teal-500" />,
        },
        {
            target: 'tour-urgencia-filter',
            title: 'Filtro por prazo',
            desc: 'Este filtro separa o que esta urgente, o que vence em breve e o que ainda esta longe.',
            icon: <AlertCircle className="w-8 h-8 text-red-500" />,
        },
        {
            target: 'tour-import-btn',
            title: 'Importar planilha',
            desc: 'Use este botao quando chegar um arquivo novo para atualizar os dados do painel.',
            icon: <UploadCloud className="w-8 h-8 text-indigo-500" />,
        },
        {
            target: 'tour-export-btn',
            title: 'Exportar arquivo',
            desc: 'Aqui voce baixa a versao atual da base para compartilhar ou guardar uma copia.',
            icon: <Download className="w-8 h-8 text-slate-700" />,
        },
        {
            target: 'tour-sync-btn',
            title: 'Atualizar da planilha online',
            desc: 'Este botao busca os dados do Google Sheets e traz a versao mais recente para o sistema.',
            icon: <TableProperties className="w-8 h-8 text-green-600" />,
        },
        {
            target: 'tour-table',
            title: 'Tabela de acompanhamento',
            desc: 'Esta e a tela principal. Aqui voce acompanha status, candidato, prazo e abre a ficha completa na ultima coluna.',
            icon: <Edit2 className="w-8 h-8 text-blue-500" />,
        },
    ],
    DASHBOARD: [
        {
            target: 'tour-tab-dashboard',
            title: 'Tela de graficos',
            desc: 'Ao clicar aqui, voce abre a area com os indicadores visuais da operacao.',
            icon: <BarChart2 className="w-8 h-8 text-indigo-500" />,
        },
        {
            target: 'tour-dashboard-panel',
            title: 'Resumo visual',
            desc: 'Nesta tela voce bate o olho nos totais, nos status e nos principais motivos em grafico.',
            icon: <LayoutDashboard className="w-8 h-8 text-sky-500" />,
        },
        {
            target: 'tour-create-chart-btn',
            title: 'Criar grafico',
            desc: 'Use este botao para montar um grafico novo e salvar no painel.',
            icon: <PlusCircle className="w-8 h-8 text-blue-500" />,
        },
    ],
    SHEETS: [
        {
            target: 'tour-tab-sheets',
            title: 'Modo planilha',
            desc: 'Esta aba abre uma visao mais parecida com planilha, boa para editar linha por linha.',
            icon: <TableProperties className="w-8 h-8 text-green-500" />,
        },
        {
            target: 'tour-sheets-table',
            title: 'Planilha editavel',
            desc: 'Preencha principalmente status, nome, candidato, contato, municipio e observacoes.',
            icon: <Check className="w-8 h-8 text-emerald-500" />,
        },
    ],
};

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
                    SISTEMA DE GESTAO DE ADMISSAO
                </h2>
            </div>
        </div>
    );
};

const WalkthroughTour = ({ section, steps, onComplete }) => {
    const [step, setStep] = useState(0);
    const [rect, setRect] = useState(null);
    const [dontShowAgain, setDontShowAgain] = useState(false);
    const [dialogSize, setDialogSize] = useState({ width: 360, height: 320 });
    const dialogRef = useRef(null);
    const currentStep = steps[step];

    useEffect(() => {
        setStep(0);
        setRect(null);
        setDontShowAgain(false);
    }, [section]);

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
        const fallbackWidth = Math.min(360, viewportWidth - (viewportPadding * 2));
        const dialogWidth = Math.min(dialogSize.width || fallbackWidth, fallbackWidth);
        const dialogHeight = Math.min(dialogSize.height || 320, viewportHeight - (viewportPadding * 2));

        if (!rect) {
            return {
                width: fallbackWidth,
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
            width: fallbackWidth,
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
                className="absolute bg-white rounded-3xl p-5 sm:p-6 shadow-[0_0_40px_rgba(0,0,0,0.3)] transition-all duration-500 animate-in zoom-in-95 overflow-y-auto"
                style={dialogStyle}
            >
                <button onClick={() => onComplete(false)} className="absolute top-4 right-4 text-slate-400 hover:text-slate-800 transition-colors" type="button"><X className="w-5 h-5" /></button>

                <div className="flex items-center gap-3 mb-4">
                    <div className="bg-indigo-50 p-2 rounded-2xl">{currentStep.icon}</div>
                    <h3 className="text-xl font-bold text-slate-800 pr-8">{currentStep.title}</h3>
                </div>

                <p className="text-slate-600 text-sm mb-8 leading-relaxed">{currentStep.desc}</p>

                <div className="flex items-center justify-between gap-3 mt-auto">
                    <div className="flex gap-1.5">
                        {steps.map((_, i) => <div key={i} className={`h-1.5 rounded-full transition-all duration-500 ${i === step ? 'w-5 bg-indigo-600' : 'w-1.5 bg-slate-200'}`} />)}
                    </div>
                    <div className="flex items-center gap-2">
                        {step > 0 && (
                            <button
                                onClick={() => setStep((s) => Math.max(0, s - 1))}
                                className="bg-slate-100 hover:bg-slate-200 text-slate-700 w-10 h-10 rounded-full flex items-center justify-center shadow-sm active:scale-95 transition-all shrink-0"
                                type="button"
                                aria-label="Voltar um passo"
                            >
                                <ChevronLeft className="w-4 h-4" />
                            </button>
                        )}
                        <button
                            onClick={() => {
                                if (step < steps.length - 1) setStep((s) => s + 1);
                                else onComplete(dontShowAgain);
                            }}
                            className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-2.5 px-5 rounded-full flex items-center gap-2 text-sm shadow-md active:scale-95 transition-all shrink-0"
                            type="button"
                        >
                            {step < steps.length - 1 ? 'Avancar' : 'Concluir'}
                            <ChevronRight className="w-4 h-4" />
                        </button>
                    </div>
                </div>

                {step === steps.length - 1 && (
                    <label className="flex items-center gap-2 mt-5 cursor-pointer group">
                        <input type="checkbox" checked={dontShowAgain} onChange={(e) => setDontShowAgain(e.target.checked)} className="rounded text-indigo-600 focus:ring-indigo-500 w-4 h-4 cursor-pointer" />
                        <span className="text-[11px] font-medium text-slate-400 group-hover:text-slate-600 transition-colors">Nao exibir este tutorial novamente</span>
                    </label>
                )}
            </div>
        </div>
    );
};

const LoginOverlay = ({ account, onCreateAccount, onLogin }) => {
    const [username, setUsername] = useState('');
    const [password, setPassword] = useState('');
    const [accessKey, setAccessKey] = useState('');
    const [loginUser, setLoginUser] = useState('');
    const [loginPass, setLoginPass] = useState('');
    const [error, setError] = useState('');

    const isFirstAccess = !account || !account.username;

    useEffect(() => {
        if (!isFirstAccess && account?.username) {
            setLoginUser(account.username);
        }
    }, [account, isFirstAccess]);

    const handleCreate = (e) => {
        e.preventDefault();
        if (!username.trim() || !password.trim()) {
            setError('Preencha usuario e senha para criar o acesso.');
            return;
        }
        if (accessKey !== 'Plansul@2025') {
            setError('Palavra-chave invalida. O sistema nao foi criado.');
            return;
        }
        onCreateAccount({ username: username.trim(), password });
    };

    const handleLogin = (e) => {
        e.preventDefault();
        if (!loginUser.trim() || !loginPass.trim()) {
            setError('Digite usuario e senha para entrar.');
            return;
        }
        const isUserValid = matchesSavedUsername(loginUser, account?.username);
        const isPasswordValid = typeof account?.password === 'string' && loginPass === account.password;
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
                <h2 className="text-2xl font-black text-slate-800 mb-2">{isFirstAccess ? 'Criar acesso' : 'Entrar no sistema'}</h2>
                <p className="text-sm text-slate-500 mb-6">{isFirstAccess ? 'Defina seu login interno. A palavra-chave de autorizacao e obrigatoria.' : 'Use o usuario e senha cadastrados para liberar o painel.'}</p>
                {!isFirstAccess && account?.username && <p className="text-xs text-slate-400 mb-4">Usuario salvo: <span className="font-bold text-slate-600">{account.username}</span>. Voce pode entrar com o nome completo ou apenas o primeiro nome.</p>}

                {isFirstAccess ? (
                    <form className="space-y-4" onSubmit={handleCreate}>
                        <input value={username} onChange={(e) => { setUsername(e.target.value); if (error) setError(''); }} className="w-full px-4 py-2.5 rounded-xl border border-slate-300 focus:ring-2 focus:ring-indigo-500" placeholder="Usuario" autoComplete="username" />
                        <input type="password" value={password} onChange={(e) => { setPassword(e.target.value); if (error) setError(''); }} className="w-full px-4 py-2.5 rounded-xl border border-slate-300 focus:ring-2 focus:ring-indigo-500" placeholder="Senha" autoComplete="new-password" />
                        <input type="password" value={accessKey} onChange={(e) => { setAccessKey(e.target.value); if (error) setError(''); }} className="w-full px-4 py-2.5 rounded-xl border border-slate-300 focus:ring-2 focus:ring-indigo-500" placeholder="Palavra-chave de autorizacao" />
                        <button type="submit" className="w-full bg-indigo-600 hover:bg-indigo-700 text-white py-2.5 rounded-xl font-bold transition-colors">Criar e entrar</button>
                    </form>
                ) : (
                    <form className="space-y-4" onSubmit={handleLogin}>
                        <input value={loginUser} onChange={(e) => { setLoginUser(e.target.value); if (error) setError(''); }} className="w-full px-4 py-2.5 rounded-xl border border-slate-300 focus:ring-2 focus:ring-indigo-500" placeholder="Usuario" autoComplete="username" />
                        <input type="password" value={loginPass} onChange={(e) => { setLoginPass(e.target.value); if (error) setError(''); }} className="w-full px-4 py-2.5 rounded-xl border border-slate-300 focus:ring-2 focus:ring-indigo-500" placeholder="Senha" autoComplete="current-password" />
                        <button type="submit" className="w-full bg-indigo-600 hover:bg-indigo-700 text-white py-2.5 rounded-xl font-bold transition-colors">Entrar</button>
                    </form>
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
    const [customChartsRaw, setCustomCharts] = useLocalStorage('vagas_custom_charts', []);
    const customCharts = Array.isArray(customChartsRaw) ? customChartsRaw : [];

    const [savedAccountRaw, setSavedAccount] = useLocalStorage('vagas_internal_account', null);
    const savedAccount = savedAccountRaw && typeof savedAccountRaw === 'object' ? savedAccountRaw : null;
    const [isAuthenticated, setIsAuthenticated] = useState(false);
    const [currentUsername, setCurrentUsername] = useState('');

    const [showWelcome, setShowWelcome] = useState(false);
    const [savedTutorialProgressRaw, setSavedTutorialProgress] = useLocalStorage('vagas_tutorial_sections', DEFAULT_TUTORIAL_PROGRESS);
    const savedTutorialProgress = useMemo(() => normalizeTutorialProgress(savedTutorialProgressRaw), [savedTutorialProgressRaw]);
    const [sessionTutorialProgress, setSessionTutorialProgress] = useState(DEFAULT_TUTORIAL_PROGRESS);
    const [showTutorial, setShowTutorial] = useState(false);
    const [tutorialSection, setTutorialSection] = useState('TABELA');
    const [pendingTutorialSection, setPendingTutorialSection] = useState(null);

    const [loading, setLoading] = useState(false);
    const [isInitialSyncing, setIsInitialSyncing] = useState(false);
    const [activeTab, setActiveTab] = useState('TABELA');
    const [searchTerm, setSearchTerm] = useState('');
    const [filters, setFilters] = useState({ status: 'TODOS', municipio: 'TODOS', motivo: 'TODOS', urgencia: 'TODOS' });
    const [showErrorsOnly, setShowErrorsOnly] = useState(false);

    const [selectedRecord, setSelectedRecord] = useState(null);
    const [isEditing, setIsEditing] = useState(false);
    const [editFormData, setEditFormData] = useState({});
    const [isChartModalOpen, setIsChartModalOpen] = useState(false);
    const [newChartData, setNewChartData] = useState({ title: '', type: 'bar', groupBy: 'CARGO' });
    const [isGSheetsModalOpen, setIsGSheetsModalOpen] = useState(false);

    const [isCinematic, setIsCinematic] = useState(false);

    const mainScrollRef = useRef(null);

    const data = Array.isArray(history.present) ? history.present : [];
    const tutorialSteps = TUTORIAL_STEPS[tutorialSection] || [];
    const hasCompletedTutorialSection = (section) => savedTutorialProgress[section] || sessionTutorialProgress[section];

    useEffect(() => {
        if (showWelcome || !isAuthenticated || data.length === 0 || showTutorial || activeTab !== 'TABELA' || hasCompletedTutorialSection('TABELA')) {
            return undefined;
        }

        const timer = setTimeout(() => {
            setTutorialSection('TABELA');
            setShowTutorial(true);
        }, 250);

        return () => clearTimeout(timer);
    }, [activeTab, data.length, isAuthenticated, savedTutorialProgress, sessionTutorialProgress, showTutorial, showWelcome]);

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

    const handleTutorialComplete = (section, dontShow) => {
        if (dontShow) {
            setSavedTutorialProgress((previous) => ({
                ...normalizeTutorialProgress(previous),
                [section]: true,
            }));
        }

        setSessionTutorialProgress((previous) => ({
            ...previous,
            [section]: true,
        }));

        setPendingTutorialSection(null);
        setShowTutorial(false);
    };

    const handleCreateAccount = (account) => {
        setSavedAccount(account);
        setCurrentUsername(account.username);
        setIsAuthenticated(true);
        setShowWelcome(true);
    };

    const handleLoginSuccess = () => {
        setCurrentUsername(savedAccount?.username || '');
        setIsAuthenticated(true);
        setShowWelcome(true);
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
        if (!(typeof window !== 'undefined' && window.XLSX)) {
            const script = document.createElement('script');
            script.src = 'https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js';
            script.async = true;
            document.body.appendChild(script);
        }

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
            if (mainScrollRef.current) mainScrollRef.current.scrollTo({ top: 0, behavior: 'auto' });
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


    const validateData = (rows) => rows.map((row) => ({ ...row, _isInvalid: !row['Nome Subs'] || !row.Status || String(row['Nome Subs']).trim() === '' }));

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
            const workbook = window.XLSX.read(buffer, { type: 'array' });
            const worksheet = workbook.Sheets.PAINEL || workbook.Sheets[workbook.SheetNames[0]];
            processDataImport(window.XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false }), false);
        } catch (error) {
            alert('Erro de parsing.');
        } finally {
            setLoading(false);
            event.target.value = null;
        }
    };

    const handleExportExcel = () => {
        if (!window.XLSX || data.length === 0) return;
        const exportData = data.map((row) => {
            const newRow = { ...row };
            delete newRow._id;
            delete newRow._isInvalid;
            return newRow;
        });
        const wb = window.XLSX.utils.book_new();
        window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.json_to_sheet(exportData), 'PAINEL_ATUALIZADO');
        window.XLSX.writeFile(wb, `Vagas_${new Date().toISOString().slice(0, 10)}.xlsx`);
    };

    const handleInlineEdit = (id, field, value) => {
        setAppData((prevData) => validateData((Array.isArray(prevData) ? prevData : []).map((item) => (item._id === id ? { ...item, [field]: value } : item))));
    };

    const handleEditClick = () => {
        setEditFormData({ ...selectedRecord });
        setIsEditing(true);
    };

    const handleSaveEdit = () => {
        setAppData((prev) => validateData((Array.isArray(prev) ? prev : []).map((item) => (item._id === editFormData._id ? editFormData : item))));
        setSelectedRecord(editFormData);
        setIsEditing(false);
    };

    const addCustomChart = () => {
        if (newChartData.title && newChartData.groupBy) {
            setCustomCharts([...customCharts, { id: Date.now(), ...newChartData }]);
            setIsChartModalOpen(false);
            setNewChartData({ title: '', type: 'bar', groupBy: 'CARGO' });
        }
    };

    const removeCustomChart = (chartId) => {
        setCustomCharts((prev) => (Array.isArray(prev) ? prev.filter((chart) => chart.id !== chartId) : []));
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
                const dias = getDaysDiff(item['Inicio Situacao']);
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
                const d_A = getDaysDiff(a['Inicio Situacao']);
                const d_B = getDaysDiff(b['Inicio Situacao']);
                if (d_A !== null && d_B !== null) return d_A - d_B;
                if (d_A !== null && d_B === null) return -1;
                if (d_A === null && d_B !== null) return 1;
            }
            return 0;
        });
    }, [data, searchTerm, filters, showErrorsOnly]);

    const tableAnimationKey = `${activeTab}-${searchTerm}-${filters.status}-${filters.municipio}-${filters.urgencia}-${showErrorsOnly}`;

    return (
        <React.Fragment>
            {!isAuthenticated && <LoginOverlay account={savedAccount} onCreateAccount={handleCreateAccount} onLogin={handleLoginSuccess} />}
            {showWelcome && isAuthenticated && <WelcomeOverlay userName={currentUsername} onFinish={() => setShowWelcome(false)} />}

            <div className={`h-screen w-full bg-slate-50 flex flex-col relative overflow-hidden transition-opacity duration-500 ease-in-out ${(showWelcome || !isAuthenticated) ? 'opacity-0' : 'opacity-100'}`}>
                <style>{`
          @keyframes cascadeSlide { 0% { opacity: 0; transform: translateY(15px); } 100% { opacity: 1; transform: translateY(0); } }
          .anim-cascade { animation: cascadeSlide 0.35s ease-out forwards; opacity: 0; }
          .cinematic-effect { filter: blur(8px) grayscale(10%) brightness(0.9); transform: scale(0.98); opacity: 0.5; pointer-events: none; }
        `}</style>

                <header className="bg-white border-b border-slate-200 z-20 shadow-sm shrink-0">
                    <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
                        <div className="flex items-center gap-3">
                            <div className="bg-indigo-600 p-2 rounded-lg shadow-sm"><Briefcase className="w-5 h-5 text-white" /></div>
                            <h1 className="text-xl font-bold text-slate-800 tracking-tight hidden lg:block">Sistema de Gestao</h1>
                        </div>

                        <div id="tour-tabs" className="flex bg-slate-100 p-1 rounded-lg overflow-x-auto hide-scrollbar relative">
                            <button onClick={() => handleTabChange('TABELA')} className={`flex items-center gap-2 px-3 sm:px-4 py-1.5 rounded-md text-sm font-semibold transition-all duration-300 whitespace-nowrap ${activeTab === 'TABELA' ? 'bg-white text-indigo-700 shadow-sm transform scale-100' : 'text-slate-500 hover:text-slate-700 scale-95'}`} type="button"><ListTodo className="w-4 h-4" /> <span className="hidden sm:inline">Acompanhamento</span>{metrics?.invalidCount > 0 && <span className="w-2 h-2 rounded-full bg-red-500 ml-1 animate-pulse" />}</button>
                            <button id="tour-tab-dashboard" onClick={() => handleTabChange('DASHBOARD', { openTutorial: true })} className={`flex items-center gap-2 px-3 sm:px-4 py-1.5 rounded-md text-sm font-semibold transition-all duration-300 whitespace-nowrap ${activeTab === 'DASHBOARD' ? 'bg-white text-indigo-700 shadow-sm transform scale-100' : 'text-slate-500 hover:text-slate-700 scale-95'}`} type="button"><LayoutDashboard className="w-4 h-4" /> <span className="hidden sm:inline">Graficos</span></button>
                            <button id="tour-tab-sheets" onClick={() => handleTabChange('SHEETS', { openTutorial: true })} className={`flex items-center gap-2 px-3 sm:px-4 py-1.5 rounded-md text-sm font-semibold transition-all duration-300 whitespace-nowrap ${activeTab === 'SHEETS' ? 'bg-green-600 text-white shadow-sm transform scale-100' : 'text-green-700 hover:bg-green-50 scale-95'}`} type="button"><TableProperties className="w-4 h-4" /> <span className="hidden sm:inline">Modo Planilha</span></button>
                        </div>

                        <div className="flex items-center gap-2 sm:gap-3">
                            <div className="flex items-center bg-slate-100 rounded-lg p-1 mr-2 border border-slate-200">
                                <button onClick={handleUndo} disabled={history.past.length === 0} className="p-1.5 text-slate-600 hover:bg-white hover:shadow-sm rounded-md disabled:opacity-30 disabled:hover:bg-transparent transition-all" type="button"><Undo2 className="w-4 h-4" /></button>
                                <div className="w-px h-4 bg-slate-300 mx-1" />
                                <button onClick={handleRedo} disabled={history.future.length === 0} className="p-1.5 text-slate-600 hover:bg-white hover:shadow-sm rounded-md disabled:opacity-30 disabled:hover:bg-transparent transition-all" type="button"><Redo2 className="w-4 h-4" /></button>
                            </div>
                            <label id="tour-import-btn" className="cursor-pointer flex items-center gap-2 px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-sm font-bold transition-all shadow-sm active:scale-95 border border-slate-200">
                                <UploadCloud className="w-4 h-4" /> <span className="hidden xl:inline">Importar</span>
                                <input type="file" accept=".xlsx, .xls" className="hidden" onChange={handleFileUpload} disabled={loading} />
                            </label>
                            <button id="tour-export-btn" onClick={handleExportExcel} className="flex items-center gap-2 px-3 py-2 bg-slate-800 text-white hover:bg-slate-900 rounded-lg text-sm font-semibold transition-all shadow-md hover:shadow-lg whitespace-nowrap active:scale-95" type="button"><Download className="w-4 h-4" /> <span className="hidden xl:inline">Exportar Validado</span></button>
                            <button id="tour-sync-btn" onClick={() => setIsGSheetsModalOpen(true)} className="flex items-center gap-2 px-3 py-2 bg-green-100 text-green-800 hover:bg-green-200 border border-green-300 rounded-lg text-sm font-bold transition-all shadow-sm whitespace-nowrap active:scale-95" type="button"><BarChart2 className="w-4 h-4" /> <span className="hidden xl:inline">Atualizar Planilha</span></button>
                        </div>
                    </div>
                </header>

                <main ref={mainScrollRef} className={`flex-1 overflow-y-auto overflow-x-hidden relative w-full bg-slate-50 transition-all duration-400 ease-in-out ${isCinematic ? 'cinematic-effect' : ''}`}>
                    <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 w-full relative z-10 min-h-full flex flex-col">
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
                                        <div className="text-xs font-bold text-indigo-600 bg-white px-2 py-1 rounded-md shadow-sm transition-all">Exibindo {filteredData.length} registros</div>
                                    </div>

                                    <div className="grid grid-cols-1 md:grid-cols-5 gap-3">
                                        <div id="tour-search" className="relative group"><Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 group-focus-within:text-indigo-500 transition-colors" /><input type="text" placeholder="Buscar..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="pl-9 pr-4 py-2.5 border border-slate-300 rounded-lg text-sm w-full focus:ring-2 focus:ring-indigo-500 transition-all" /></div>
                                        <select id="tour-status-filter" value={filters.status} onChange={(e) => setFilters((f) => ({ ...f, status: e.target.value }))} className="px-3 py-2.5 border border-slate-300 rounded-lg text-sm focus:ring-2 focus:ring-indigo-500 bg-white font-medium transition-all"><option value="TODOS">Todos os Status</option>{listOptions.status.map((s) => <option key={s} value={s}>{s}</option>)}</select>
                                        <select id="tour-municipio-filter" value={filters.municipio} onChange={(e) => setFilters((f) => ({ ...f, municipio: e.target.value }))} className="px-3 py-2.5 border border-slate-300 rounded-lg text-sm focus:ring-2 focus:ring-indigo-500 bg-white transition-all"><option value="TODOS">Municipios</option>{listOptions.municipios.map((m) => <option key={m} value={m}>{m}</option>)}</select>
                                        <select id="tour-urgencia-filter" value={filters.urgencia} onChange={(e) => setFilters((f) => ({ ...f, urgencia: e.target.value }))} className="px-3 py-2.5 border border-indigo-300 rounded-lg text-sm focus:ring-2 focus:ring-indigo-500 bg-indigo-50 font-bold text-indigo-900 transition-all"><option value="TODOS">Filtro Termico</option><option value="URGENTE">Urgente (&lt;= 5d)</option><option value="MEDIA">Em breve (6 a 30d)</option><option value="LONGE">Longo Prazo</option></select>
                                        {showErrorsOnly ? (
                                            <button id="tour-errors-filter" onClick={() => setShowErrorsOnly(false)} className="bg-slate-200 text-slate-700 font-bold text-xs rounded-lg hover:bg-slate-300 transition-all flex items-center justify-center gap-2 active:scale-95" type="button"><X className="w-4 h-4" /> Mostrar Tudo</button>
                                        ) : metrics?.invalidCount > 0 && (
                                            <button id="tour-errors-filter" onClick={() => setShowErrorsOnly(true)} className="bg-red-100 text-red-700 font-bold text-xs rounded-lg hover:bg-red-200 transition-all flex items-center justify-center gap-2 active:scale-95 animate-pulse" type="button"><AlertCircle className="w-4 h-4" /> Corrigir Erros</button>
                                        )}
                                    </div>
                                </div>

                                <div className="overflow-x-auto relative flex-1 min-h-[400px]">
                                    <table className="w-full text-left text-sm whitespace-nowrap">
                                        <thead className="bg-slate-800 border-b border-slate-700 text-slate-200 font-semibold text-xs uppercase tracking-wider sticky top-0 z-10">
                                            <tr><th className="px-6 py-4">Status Rapido</th><th className="px-6 py-4">Vaga</th><th className="px-6 py-4">Candidato</th><th className="px-6 py-4">Prazo Termico</th><th className="px-6 py-4 text-right">Ficha</th></tr>
                                        </thead>
                                        <tbody id="tour-table" key={tableAnimationKey} className="divide-y divide-slate-200/50">
                                            {filteredData.slice(0, 100).map((row, index) => {
                                                const status = safeGet(row, 'Status');
                                                const candidato = row.Candidato || 'SEM COBERTURA';
                                                const diasParaInicio = getDaysDiff(row['Inicio Situacao']);
                                                return (
                                                    <tr key={row._id} className={`${getRowThermalClass(diasParaInicio, status, candidato, row._isInvalid)} group anim-cascade transition-all duration-300 hover:shadow-sm`} style={{ animationDelay: `${index * 15}ms` }}>
                                                        <td className="px-6 py-4 relative">
                                                            {row._isInvalid && <AlertCircle className="absolute left-1 top-1/2 -translate-y-1/2 w-4 h-4 text-red-600 animate-pulse" />}
                                                            <select value={status} onChange={(e) => handleInlineEdit(row._id, 'Status', e.target.value)} className={`appearance-none w-32 ml-2 px-2 py-1.5 rounded-lg text-xs font-bold cursor-pointer focus:ring-2 focus:ring-indigo-500 shadow-sm transition-all ${STATUS_COLORS[status] || 'bg-slate-100 text-slate-800 border-slate-200'}`}>
                                                                <option value="ABERTA">ABERTA</option><option value="FECHADA">FECHADA</option><option value="ENCAMINHADA">ENCAMINHADA</option><option value="CANCELADA">CANCELADA</option><option value="PAUSADA">PAUSADA</option>
                                                            </select>
                                                        </td>
                                                        <td className="px-6 py-4"><div className={`font-bold ${row._isInvalid ? 'text-red-700' : 'text-slate-900'}`}>{row['Nome Subs'] || 'FALTANDO'}</div><div className="text-slate-600 text-xs mt-0.5">{row.CARGO} • {row['NRE / MUNICIPIO']}</div></td>
                                                        <td className="px-6 py-4"><div className={`font-bold ${candidato === 'SEM COBERTURA' ? 'text-red-700' : 'text-slate-800'}`}>{candidato}</div><div className="text-slate-500 text-xs mt-0.5">{row['Contato Candidato'] || 'Sem contato'}</div></td>
                                                        <td className="px-6 py-4"><div className="font-bold text-slate-800">{row['Inicio Situacao'] || '-'}</div>
                                                            {diasParaInicio !== null && !['FECHADA', 'ENCAMINHADA', 'CANCELADA'].includes(status) && candidato === 'SEM COBERTURA' && (
                                                                <div className={`text-xs mt-1 font-bold inline-block px-2 py-0.5 rounded shadow-sm transition-transform group-hover:scale-105 ${diasParaInicio < 0 ? 'bg-red-600 text-white' : diasParaInicio === 0 ? 'bg-red-500 text-white' : diasParaInicio <= 5 ? 'bg-orange-500 text-white' : diasParaInicio <= 15 ? 'bg-yellow-400 text-yellow-900' : diasParaInicio <= 30 ? 'bg-lime-500 text-lime-900' : 'bg-green-500 text-white'}`}>
                                                                    {diasParaInicio < 0 ? `Atrasado ${Math.abs(diasParaInicio)}d` : diasParaInicio === 0 ? 'E Hoje!' : `Faltam ${diasParaInicio}d`}
                                                                </div>
                                                            )}
                                                        </td>
                                                        <td className="px-6 py-4 text-right"><button onClick={() => { setSelectedRecord(row); setIsEditing(false); }} className="p-2.5 bg-white hover:bg-indigo-50 rounded-lg shadow-sm border border-slate-200 transition-all hover:border-indigo-300 hover:shadow-md active:scale-95" type="button"><Edit2 className="w-4 h-4 text-slate-600 hover:text-indigo-700" /></button></td>
                                                    </tr>
                                                );
                                            })}
                                        </tbody>
                                    </table>
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
                                <div className="bg-green-600 text-white p-3 flex justify-between shadow-sm z-10 shrink-0"><div className="flex items-center gap-2"><TableProperties className="w-5 h-5 text-green-100" /><h2 className="font-bold">Planilha editavel</h2></div></div>
                                <div className="overflow-auto flex-1 bg-slate-100 p-2">
                                    <table className="w-full text-left text-sm border-collapse bg-white shadow-sm ring-1 ring-slate-200">
                                        <thead className="bg-slate-100 border-b-2 border-slate-300 text-slate-700 font-bold text-xs sticky top-0 z-10 shadow-sm">
                                            <tr><th className="px-3 py-2 border-r min-w-[120px]">Status</th><th className="px-3 py-2 border-r min-w-[200px]">Profissional Saida</th><th className="px-3 py-2 border-r bg-green-50 min-w-[200px]">Candidato Cobertura</th><th className="px-3 py-2 border-r bg-green-50 min-w-[120px]">Telefone</th><th className="px-3 py-2 border-r min-w-[150px]">Municipio</th><th className="px-3 py-2 min-w-[200px]">Observacoes</th></tr>
                                        </thead>
                                        <tbody key={`sheets-${tableAnimationKey}`} className="divide-y divide-slate-200 font-medium">
                                            {filteredData.slice(0, 50).map((row, i) => (
                                                <tr key={row._id} className={`hover:bg-blue-50/50 transition-colors anim-cascade ${row._isInvalid ? 'bg-red-50' : ''}`} style={{ animationDelay: `${i * 15}ms` }}>
                                                    <td className="border-r p-0"><select value={row.Status || ''} onChange={(e) => handleInlineEdit(row._id, 'Status', e.target.value)} className="w-full h-full px-3 py-2 appearance-none cursor-pointer focus:bg-blue-100 bg-transparent transition-colors"><option value="ABERTA">ABERTA</option><option value="FECHADA">FECHADA</option><option value="ENCAMINHADA">ENCAMINHADA</option><option value="CANCELADA">CANCELADA</option><option value="PAUSADA">PAUSADA</option></select></td>
                                                    <td className={`border-r p-0 transition-colors ${row._isInvalid && !row['Nome Subs'] ? 'ring-2 ring-inset ring-red-500 bg-red-50' : ''}`}><input type="text" value={row['Nome Subs'] || ''} onChange={(e) => handleInlineEdit(row._id, 'Nome Subs', e.target.value)} className="w-full h-full px-3 py-2 focus:bg-blue-100 bg-transparent transition-colors placeholder:text-red-300" placeholder="Obrigatorio" /></td>
                                                    <td className="border-r p-0 bg-green-50/30"><input type="text" value={row.Candidato || ''} onChange={(e) => handleInlineEdit(row._id, 'Candidato', e.target.value)} className="w-full h-full px-3 py-2 focus:bg-green-100 font-bold text-green-900 bg-transparent transition-colors placeholder:text-green-300" placeholder="Sem Candidato" /></td>
                                                    <td className="border-r p-0 bg-green-50/30"><input type="text" value={row['Contato Candidato'] || ''} onChange={(e) => handleInlineEdit(row._id, 'Contato Candidato', e.target.value)} className="w-full h-full px-3 py-2 focus:bg-green-100 bg-transparent transition-colors" /></td>
                                                    <td className="border-r p-0"><input type="text" value={row['NRE / MUNICIPIO'] || ''} onChange={(e) => handleInlineEdit(row._id, 'NRE / MUNICIPIO', e.target.value)} className="w-full h-full px-3 py-2 focus:bg-blue-100 bg-transparent transition-colors" /></td>
                                                    <td className="p-0"><input type="text" value={row['OBS:'] || ''} onChange={(e) => handleInlineEdit(row._id, 'OBS:', e.target.value)} className="w-full h-full px-3 py-2 focus:bg-blue-100 bg-transparent text-xs transition-colors" placeholder="Adicionar nota..." /></td>
                                                </tr>
                                            ))}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        )}
                    </div>
                </main>
            </div>

            {showTutorial && !showWelcome && tutorialSteps.length > 0 && (
                <WalkthroughTour
                    key={tutorialSection}
                    section={tutorialSection}
                    steps={tutorialSteps}
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
                    <div className="bg-white rounded-3xl shadow-2xl w-full max-w-md p-8 animate-in zoom-in-95 duration-300">
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

            {selectedRecord && (
                <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[90] animate-in fade-in duration-200">
                    <div className="bg-white rounded-3xl shadow-2xl w-full max-w-3xl max-h-[90vh] flex flex-col overflow-hidden animate-in zoom-in-95 duration-300">
                        <div className="px-8 py-5 border-b flex justify-between bg-white items-center"><h3 className="text-xl font-bold text-slate-800 flex gap-2 items-center"><Briefcase className="w-6 h-6 text-indigo-600" />{isEditing ? 'Configurando Matriz' : 'Ficha Completa'}</h3><div className="flex gap-2 items-center">{!isEditing ? <button onClick={handleEditClick} className="flex items-center gap-1.5 px-4 py-2 bg-indigo-50 text-indigo-700 font-bold rounded-xl text-sm hover:bg-indigo-100 transition-colors" type="button"><Edit2 className="w-4 h-4" /> Editar Todos os Campos</button> : <button onClick={handleSaveEdit} className="flex items-center gap-1.5 px-5 py-2 bg-green-600 text-white font-bold rounded-xl text-sm hover:bg-green-700 shadow-md transition-all active:scale-95" type="button"><Save className="w-4 h-4" /> Salvar Edicoes</button>}<button onClick={() => setSelectedRecord(null)} className="p-2 text-slate-400 hover:bg-slate-100 rounded-xl ml-2 transition-colors" type="button"><X className="w-5 h-5" /></button></div></div>
                        <div className="p-8 overflow-y-auto flex-1 bg-slate-50 relative">
                            {!isEditing ? (
                                <div className="space-y-6"><div className="grid grid-cols-1 md:grid-cols-2 gap-8"><div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 space-y-4"><h4 className="text-xs font-bold text-slate-400 uppercase border-b pb-2">Profissional</h4><p className="text-sm"><b>Nome:</b> {selectedRecord['Nome Subs']}</p><p className="text-sm"><b>Cargo:</b> {selectedRecord.CARGO}</p></div><div className="bg-indigo-50 p-6 rounded-2xl border border-indigo-100 shadow-sm space-y-4"><h4 className="text-xs font-bold text-slate-400 uppercase border-b border-indigo-200 pb-2">Candidato e Cobertura</h4><p className="text-sm"><b>Status:</b> {selectedRecord.Situacao}</p><p className="text-sm"><b>Candidato:</b> {selectedRecord.Candidato}</p></div></div></div>
                            ) : (
                                <div className="space-y-6 bg-white p-8 rounded-2xl border shadow-inner"><div className="grid grid-cols-1 md:grid-cols-2 gap-6"><div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase">Status</label><select value={editFormData.Status || ''} onChange={(e) => setEditFormData({ ...editFormData, Status: e.target.value })} className="w-full px-4 py-2.5 border rounded-xl focus:ring-2 focus:ring-indigo-500 bg-slate-50"><option value="ABERTA">ABERTA</option><option value="FECHADA">FECHADA</option><option value="ENCAMINHADA">ENCAMINHADA</option><option value="CANCELADA">CANCELADA</option><option value="PAUSADA">PAUSADA</option></select></div><div className="space-y-1.5"><label className="text-xs font-bold text-slate-500 uppercase">Candidato</label><input type="text" value={editFormData.Candidato || ''} onChange={(e) => setEditFormData({ ...editFormData, Candidato: e.target.value })} className="w-full px-4 py-2.5 border rounded-xl focus:ring-2 focus:ring-indigo-500 bg-slate-50" /></div></div></div>
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
