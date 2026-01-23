import { useState, useEffect, useMemo } from 'react';
import { Linkedin, TrendingUp, TrendingDown, Minus, Users, FileText, Eye, Plus, X, Download, Calendar, BarChart3, List, Loader2, Cloud, CloudOff } from 'lucide-react';
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, BarChart, Bar, Legend } from 'recharts';
import * as XLSX from 'xlsx';
import { getRegistros, addRegistro, updateRegistro, deleteRegistro } from './firebase';

const ACCOUNTS = [
  { key: 'gmc', label: 'GMC', color: '#8B5CF6' },        // Morado
  { key: 'educa', label: 'EDUCA', color: '#F97316' },    // Naranja
  { key: 'maribel', label: 'MARIBEL', color: '#10B981' }, // Verde
  { key: 'egisto', label: 'EGISTO', color: '#EC4899' },   // Rosa
];

const METRICS = [
  { key: 'seg', label: 'Seguidores', icon: Users },
  { key: 'pub', label: 'Publicaciones', icon: FileText },
  { key: 'imp', label: 'Impresiones', icon: Eye },
];

const Card = ({ children, className = '' }) => (
  <div className={`bg-white rounded-xl shadow-sm border p-4 ${className}`}>{children}</div>
);

const Button = ({ children, onClick, variant = 'primary', disabled, className = '' }) => {
  const base = 'px-4 py-2 rounded-lg font-medium transition-all flex items-center gap-2 disabled:opacity-50';
  const styles = {
    primary: 'bg-[#0077B5] text-white hover:bg-[#005582]',
    secondary: 'bg-gray-100 text-gray-700 hover:bg-gray-200',
    danger: 'bg-red-500 text-white hover:bg-red-600',
  };
  return <button onClick={onClick} disabled={disabled} className={`${base} ${styles[variant]} ${className}`}>{children}</button>;
};

const Trend = ({ current, previous }) => {
  if (!previous) return <span className="text-gray-400 text-sm">—</span>;
  const diff = current - previous;
  const pct = previous > 0 ? ((diff / previous) * 100).toFixed(1) : 0;
  if (diff > 0) return <span className="text-green-600 text-sm flex items-center gap-1"><TrendingUp size={14} />+{pct}%</span>;
  if (diff < 0) return <span className="text-red-500 text-sm flex items-center gap-1"><TrendingDown size={14} />{pct}%</span>;
  return <span className="text-gray-400 text-sm flex items-center gap-1"><Minus size={14} />0%</span>;
};

const DataModal = ({ isOpen, onClose, onSave, editData, existingDates, loading }) => {
  const [form, setForm] = useState({});
  
  useEffect(() => {
    if (editData) {
      setForm(editData);
    } else {
      const empty = { fecha: new Date().toISOString().split('T')[0] };
      ACCOUNTS.forEach(a => METRICS.forEach(m => empty[`${a.key}_${m.key}`] = 0));
      setForm(empty);
    }
  }, [editData, isOpen]);

  if (!isOpen) return null;

  const handleSubmit = () => {
    if (!form.fecha) return alert('Selecciona una fecha');
    if (!editData && existingDates.includes(form.fecha)) return alert('Ya existe un registro para esta fecha');
    onSave(form);
  };

  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-xl w-full max-w-2xl max-h-[90vh] overflow-y-auto">
        <div className="p-4 border-b flex justify-between items-center sticky top-0 bg-white">
          <h2 className="text-lg font-bold">{editData ? 'Editar' : 'Nuevo'} Registro</h2>
          <button onClick={onClose} className="p-1 hover:bg-gray-100 rounded"><X size={20} /></button>
        </div>
        <div className="p-4 space-y-4">
          <div>
            <label className="block text-sm font-medium mb-1">Fecha</label>
            <input
              type="date"
              value={form.fecha || ''}
              onChange={e => setForm({ ...form, fecha: e.target.value })}
              disabled={!!editData}
              className="w-full border rounded-lg px-3 py-2 disabled:bg-gray-100"
            />
          </div>
          
          {ACCOUNTS.map(account => (
            <div key={account.key} className="border rounded-lg p-3" style={{ borderLeftColor: account.color, borderLeftWidth: '4px' }}>
              <h3 className="font-semibold mb-2 flex items-center gap-2" style={{ color: account.color }}>
                <Linkedin size={16} /> {account.label}
              </h3>
              <div className="grid grid-cols-3 gap-3">
                {METRICS.map(m => (
                  <div key={m.key}>
                    <label className="block text-xs text-gray-500 mb-1">{m.label}</label>
                    <input
                      type="number"
                      min="0"
                      value={form[`${account.key}_${m.key}`] || 0}
                      onChange={e => setForm({ ...form, [`${account.key}_${m.key}`]: parseInt(e.target.value) || 0 })}
                      className="w-full border rounded px-2 py-1 text-sm"
                    />
                  </div>
                ))}
              </div>
            </div>
          ))}
        </div>
        <div className="p-4 border-t flex justify-end gap-2 sticky bottom-0 bg-white">
          <Button variant="secondary" onClick={onClose}>Cancelar</Button>
          <Button onClick={handleSubmit} disabled={loading}>
            {loading ? <><Loader2 size={16} className="animate-spin" /> Guardando...</> : (editData ? 'Guardar' : 'Agregar')}
          </Button>
        </div>
      </div>
    </div>
  );
};

const exportExcel = (data) => {
  const wb = XLSX.utils.book_new();
  const headers = ['FECHA'];
  ACCOUNTS.forEach(a => METRICS.forEach(m => headers.push(`${a.label}_${m.label.toUpperCase()}`)));
  
  const rows = data.map(r => {
    const row = [r.fecha];
    ACCOUNTS.forEach(a => METRICS.forEach(m => row.push(r[`${a.key}_${m.key}`] || 0)));
    return row;
  });
  
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  XLSX.utils.book_append_sheet(wb, ws, 'LinkedIn');
  XLSX.writeFile(wb, `Social_LinkedIn_${new Date().toISOString().split('T')[0]}.xlsx`);
};

export default function App() {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [online, setOnline] = useState(true);
  const [view, setView] = useState('dashboard');
  const [modal, setModal] = useState({ open: false, edit: null });
  const [selectedMetric, setSelectedMetric] = useState('seg');
  const [toast, setToast] = useState(null);

  // Cargar datos de Firebase al iniciar
  useEffect(() => {
    loadData();
  }, []);

  const loadData = async () => {
    setLoading(true);
    try {
      const registros = await getRegistros();
      setData(registros);
      setOnline(true);
    } catch (error) {
      console.error('Error cargando datos:', error);
      setOnline(false);
      notify('❌ Error al conectar con Firebase');
    }
    setLoading(false);
  };

  const notify = msg => { setToast(msg); setTimeout(() => setToast(null), 3000); };

  const sorted = useMemo(() => [...data].sort((a, b) => new Date(a.fecha) - new Date(b.fecha)), [data]);
  const latest = sorted[sorted.length - 1];
  const prev = sorted[sorted.length - 2];

  const chartData = useMemo(() => sorted.map(r => ({
    fecha: new Date(r.fecha).toLocaleDateString('es-MX', { day: '2-digit', month: 'short' }),
    ...r
  })), [sorted]);

  const handleSave = async (newData) => {
    setSaving(true);
    try {
      const exists = data.find(d => d.fecha === newData.fecha);
      if (exists) {
        await updateRegistro(newData.fecha, newData);
        setData(prev => prev.map(d => d.fecha === newData.fecha ? { ...newData, id: newData.fecha } : d));
        notify('✅ Actualizado en Firebase');
      } else {
        await addRegistro(newData);
        setData(prev => [...prev, { ...newData, id: newData.fecha }]);
        notify('✅ Guardado en Firebase');
      }
      setModal({ open: false, edit: null });
    } catch (error) {
      console.error('Error guardando:', error);
      notify('❌ Error al guardar');
    }
    setSaving(false);
  };

  const handleDelete = async (fecha) => {
    if (!confirm('¿Eliminar registro?')) return;
    try {
      await deleteRegistro(fecha);
      setData(prev => prev.filter(d => d.fecha !== fecha));
      notify('🗑️ Eliminado de Firebase');
    } catch (error) {
      console.error('Error eliminando:', error);
      notify('❌ Error al eliminar');
    }
  };

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center">
        <div className="text-center">
          <Loader2 size={48} className="animate-spin text-[#0077B5] mx-auto mb-4" />
          <p className="text-gray-600">Conectando con Firebase...</p>
        </div>
      </div>
    );
  }

  const currentMetric = METRICS.find(m => m.key === selectedMetric);

  return (
    <div className="min-h-screen bg-gray-50">
      {toast && (
        <div className="fixed top-4 right-4 bg-gray-900 text-white px-4 py-2 rounded-lg shadow-lg z-50">
          {toast}
        </div>
      )}

      <header className="bg-white border-b sticky top-0 z-40">
        <div className="max-w-7xl mx-auto px-4 py-3 flex justify-between items-center">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-[#0077B5] rounded-lg flex items-center justify-center">
              <Linkedin className="text-white" size={24} />
            </div>
            <div>
              <h1 className="font-bold text-gray-800">Social Tracker 360</h1>
              <p className="text-xs text-gray-500">LinkedIn Analytics</p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <div className={`flex items-center gap-1 text-xs px-2 py-1 rounded ${online ? 'bg-green-100 text-green-700' : 'bg-red-100 text-red-700'}`}>
              {online ? <Cloud size={14} /> : <CloudOff size={14} />}
              {online ? 'Conectado' : 'Sin conexión'}
            </div>
            <Button variant="secondary" onClick={loadData} disabled={loading}>
              <Loader2 size={16} className={loading ? 'animate-spin' : ''} /> Actualizar
            </Button>
            <Button onClick={() => setModal({ open: true, edit: null })}>
              <Plus size={16} /> Agregar
            </Button>
          </div>
        </div>
        
        <div className="max-w-7xl mx-auto px-4 pb-2 flex gap-1">
          {[
            { id: 'dashboard', icon: BarChart3, label: 'Dashboard' },
            { id: 'trends', icon: TrendingUp, label: 'Tendencias' },
            { id: 'history', icon: List, label: 'Historial' },
          ].map(tab => (
            <button
              key={tab.id}
              onClick={() => setView(tab.id)}
              className={`px-3 py-1.5 rounded-lg text-sm flex items-center gap-1.5 transition-all ${
                view === tab.id ? 'bg-[#0077B5] text-white' : 'text-gray-600 hover:bg-gray-100'
              }`}
            >
              <tab.icon size={14} /> {tab.label}
            </button>
          ))}
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-6">
        {view === 'dashboard' && (
          <div className="space-y-6">
            {data.length === 0 ? (
              <Card className="text-center py-12">
                <Calendar size={48} className="mx-auto text-gray-300 mb-4" />
                <p className="text-gray-500 mb-4">No hay datos todavía</p>
                <Button onClick={() => setModal({ open: true, edit: null })}>
                  <Plus size={16} /> Agregar primer registro
                </Button>
              </Card>
            ) : (
              <>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                  {ACCOUNTS.map(account => (
                    <Card key={account.key} className="hover:shadow-md transition-shadow" style={{ borderTopColor: account.color, borderTopWidth: '4px' }}>
                      <div className="flex items-center gap-2 mb-3">
                        <div className="w-8 h-8 rounded flex items-center justify-center" style={{ backgroundColor: account.color }}>
                          <Linkedin className="text-white" size={16} />
                        </div>
                        <span className="font-bold text-gray-800">{account.label}</span>
                      </div>
                      <div className="space-y-2">
                        {METRICS.map(m => {
                          const key = `${account.key}_${m.key}`;
                          const current = latest?.[key] || 0;
                          const previous = prev?.[key];
                          return (
                            <div key={m.key} className="flex justify-between items-center">
                              <div className="flex items-center gap-2 text-sm text-gray-600">
                                <m.icon size={14} />
                                {m.label}
                              </div>
                              <div className="text-right">
                                <span className="font-semibold">{current.toLocaleString()}</span>
                                <div className="ml-2 inline-block"><Trend current={current} previous={previous} /></div>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </Card>
                  ))}
                </div>

                {chartData.length > 1 && (
                  <Card>
                    <div className="flex justify-between items-center mb-4">
                      <h2 className="font-bold text-gray-800">Comparativo de Seguidores</h2>
                    </div>
                    <ResponsiveContainer width="100%" height={300}>
                      <BarChart data={[chartData[chartData.length - 1]]}>
                        <CartesianGrid strokeDasharray="3 3" />
                        <XAxis dataKey="fecha" />
                        <YAxis />
                        <Tooltip />
                        <Legend />
                        {ACCOUNTS.map(account => (
                          <Bar 
                            key={account.key} 
                            dataKey={`${account.key}_seg`} 
                            name={account.label}
                            fill={account.color}
                          />
                        ))}
                      </BarChart>
                    </ResponsiveContainer>
                  </Card>
                )}
              </>
            )}
          </div>
        )}

        {view === 'trends' && (
          <div className="space-y-6">
            <Card>
              <div className="flex flex-wrap gap-2 mb-4">
                {METRICS.map(metric => (
                  <button
                    key={metric.key}
                    onClick={() => setSelectedMetric(metric.key)}
                    className={`px-4 py-2 rounded-lg text-sm font-medium transition-all flex items-center gap-2 ${
                      selectedMetric === metric.key 
                        ? 'bg-[#0077B5] text-white' 
                        : 'bg-gray-100 text-gray-600 hover:bg-gray-200'
                    }`}
                  >
                    <metric.icon size={16} /> {metric.label}
                  </button>
                ))}
              </div>

              {/* Leyenda de colores */}
              <div className="flex flex-wrap gap-4 mb-4 pb-4 border-b">
                {ACCOUNTS.map(account => (
                  <div key={account.key} className="flex items-center gap-2">
                    <div className="w-3 h-3 rounded-full" style={{ backgroundColor: account.color }}></div>
                    <span className="text-sm font-medium">{account.label}</span>
                  </div>
                ))}
              </div>
              
              {chartData.length > 1 ? (
                <ResponsiveContainer width="100%" height={400}>
                  <LineChart data={chartData}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis dataKey="fecha" />
                    <YAxis />
                    <Tooltip />
                    <Legend />
                    {ACCOUNTS.map(account => (
                      <Line
                        key={account.key}
                        type="monotone"
                        dataKey={`${account.key}_${selectedMetric}`}
                        name={account.label}
                        stroke={account.color}
                        strokeWidth={3}
                        dot={{ r: 5, fill: account.color }}
                      />
                    ))}
                  </LineChart>
                </ResponsiveContainer>
              ) : (
                <div className="text-center py-12 text-gray-500">
                  Necesitas al menos 2 registros para ver tendencias
                </div>
              )}
            </Card>
          </div>
        )}

        {view === 'history' && (
          <Card>
            <div className="flex justify-between items-center mb-4">
              <h2 className="font-bold text-gray-800">Historial de Registros</h2>
              <Button variant="secondary" onClick={() => exportExcel(sorted)}>
                <Download size={16} /> Excel
              </Button>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b">
                    <th className="text-left py-2 px-2">Fecha</th>
                    {ACCOUNTS.map(a => (
                      <th key={a.key} colSpan={3} className="text-center py-2 px-2" style={{ backgroundColor: `${a.color}20` }}>
                        <span style={{ color: a.color }}>{a.label}</span>
                      </th>
                    ))}
                    <th className="text-center py-2 px-2">Acciones</th>
                  </tr>
                  <tr className="border-b text-xs text-gray-500">
                    <th></th>
                    {ACCOUNTS.map(a => (
                      METRICS.map(m => (
                        <th key={`${a.key}-${m.key}`} className="py-1 px-1">{m.label.substring(0, 3)}</th>
                      ))
                    ))}
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {[...sorted].reverse().map(row => (
                    <tr key={row.fecha} className="border-b hover:bg-gray-50">
                      <td className="py-2 px-2 font-medium">{row.fecha}</td>
                      {ACCOUNTS.map(a => (
                        METRICS.map(m => (
                          <td key={`${a.key}-${m.key}`} className="py-2 px-1 text-center">
                            {(row[`${a.key}_${m.key}`] || 0).toLocaleString()}
                          </td>
                        ))
                      ))}
                      <td className="py-2 px-2 text-center">
                        <button
                          onClick={() => setModal({ open: true, edit: row })}
                          className="text-blue-600 hover:underline mr-2"
                        >
                          Editar
                        </button>
                        <button
                          onClick={() => handleDelete(row.fecha)}
                          className="text-red-500 hover:underline"
                        >
                          Eliminar
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>
        )}
      </main>

      <DataModal
        isOpen={modal.open}
        onClose={() => setModal({ open: false, edit: null })}
        onSave={handleSave}
        editData={modal.edit}
        existingDates={data.map(d => d.fecha)}
        loading={saving}
      />

      <footer className="border-t bg-white mt-8">
        <div className="max-w-7xl mx-auto px-4 py-4 flex justify-between items-center">
          <div className="flex items-center gap-2">
            <Linkedin size={16} className="text-[#0077B5]" />
            <span className="font-semibold">Social Tracker 360</span>
            <span className="text-xs bg-green-100 text-green-700 px-2 py-0.5 rounded">Firebase</span>
          </div>
          <p className="text-xs text-gray-400">© 2026 - Datos en la nube ☁️</p>
        </div>
      </footer>
    </div>
  );
}
