/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef } from 'react';
import { 
  LayoutDashboard, 
  Package, 
  Upload, 
  Settings as SettingsIcon, 
  Search, 
  Filter, 
  ChevronRight, 
  CheckCircle2, 
  Clock, 
  AlertCircle,
  LogOut,
  FileText,
  Plus,
  ArrowUpRight,
  ArrowDownRight,
  Loader2,
  Trash2,
  Save,
  ExternalLink,
  Camera,
  Menu,
  X,
  Edit3,
  Download
} from 'lucide-react';
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer, 
  Cell,
  PieChart,
  Pie
} from 'recharts';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import { GoogleGenAI } from "@google/genai";
import * as XLSX from 'xlsx';

// --- Utility ---
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// --- Types ---
type ItemStatus = 'stock' | 'usado' | 'abono';

interface InvoiceItem {
  id: string;
  referencia: string;
  descripcion: string;
  cantidad: number;
  pvp: number;
  descuento: number;
  precio_final: number;
  total_item: number;
  beneficio: number;
  estado: ItemStatus;
  proveedor: string;
  fecha: string;
  numero_factura: string;
  abonoRecibido?: boolean;
}

interface ExtractionResult {
  documentos: Array<{
    proveedor: string;
    fecha: string;
    numero_factura: string;
    items: Array<{
      referencia: string;
      descripcion: string;
      cantidad: number;
      pvp: number;
      descuento: number;
      precio_final: number;
      total_item: number;
    }>;
  }>;
}

// --- Components ---

const Logo = () => (
  <div className="flex flex-col items-center justify-center py-6 px-4">
    <div className="relative group cursor-default">
      <div className="text-4xl font-black tracking-tighter italic text-transparent bg-clip-text bg-gradient-to-b from-gray-100 to-gray-500 drop-shadow-[0_2px_2px_rgba(0,0,0,0.8)]">
        GARGAM
      </div>
      <div className="text-sm font-bold tracking-[0.3em] text-[#F27D26] mt-[-8px] text-center uppercase">
        Performance
      </div>
      <div className="absolute -bottom-2 left-0 w-full h-[2px] bg-gradient-to-r from-transparent via-[#F27D26] to-transparent opacity-50"></div>
      <div className="absolute -top-2 left-0 w-full h-[1px] bg-gradient-to-r from-transparent via-gray-400 to-transparent opacity-30"></div>
    </div>
  </div>
);

const Card = ({ children, className, title, onClick }: { children: React.ReactNode, className?: string, title?: string, onClick?: () => void }) => (
  <div 
    className={cn("bg-[#151619] border border-white/5 rounded-xl overflow-hidden shadow-2xl", className)}
    onClick={onClick}
  >
    {title && (
      <div className="px-6 py-4 border-bottom border-white/5 bg-white/2">
        <h3 className="text-sm font-semibold uppercase tracking-wider text-gray-400">{title}</h3>
      </div>
    )}
    <div className="p-6">{children}</div>
  </div>
);

const StatCard = ({ title, value, icon: Icon, trend, trendValue, onClick }: { title: string, value: string, icon: any, trend?: 'up' | 'down', trendValue?: string, onClick?: () => void }) => (
  <Card 
    className={cn("relative overflow-hidden group transition-all duration-300", onClick && "cursor-pointer hover:border-[#F27D26]/50 hover:bg-white/2")}
    onClick={onClick}
  >
    <div className="flex justify-between items-start">
      <div>
        <p className="text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">{title}</p>
        <h3 className="text-2xl font-bold text-white">{value}</h3>
        {trend && (
          <div className={cn("flex items-center mt-2 text-xs font-medium", trend === 'up' ? "text-emerald-400" : "text-rose-400")}>
            {trend === 'up' ? <ArrowUpRight className="w-3 h-3 mr-1" /> : <ArrowDownRight className="w-3 h-3 mr-1" />}
            {trendValue}
          </div>
        )}
      </div>
      <div className="p-3 bg-white/5 rounded-lg group-hover:bg-[#F27D26]/20 transition-colors">
        <Icon className="w-5 h-5 text-[#F27D26]" />
      </div>
    </div>
    <div className="absolute bottom-0 left-0 w-full h-1 bg-gradient-to-r from-transparent via-[#F27D26]/20 to-transparent"></div>
  </Card>
);

export default function App() {
  const [activeTab, setActiveTab] = useState<'dashboard' | 'inventory' | 'upload' | 'settings'>('dashboard');
  const [inventory, setInventory] = useState<InvoiceItem[]>([]);
  const [isUploading, setIsUploading] = useState(false);
  const [showManualForm, setShowManualForm] = useState(false);
  const [manualItem, setManualItem] = useState({
    fecha: new Date().toISOString().split('T')[0],
    proveedor: '',
    descripcion: '',
    referencia: '',
    cantidad: 1,
    pvp: 0,
    descuento: 0
  });
  const [isSyncing, setIsSyncing] = useState(false);
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const [tokens, setTokens] = useState<any>(null);
  const [spreadsheetId, setSpreadsheetId] = useState(import.meta.env.VITE_GOOGLE_SPREADSHEET_ID || '');
  const [searchTerm, setSearchTerm] = useState('');
  const [filterStatus, setFilterStatus] = useState<ItemStatus | 'all'>('all');
  const socketRef = useRef<WebSocket | null>(null);
  const [deferredPrompt, setDeferredPrompt] = useState<any>(null);

  // PWA Install Prompt
  useEffect(() => {
    const handleBeforeInstallPrompt = (e: Event) => {
      e.preventDefault();
      setDeferredPrompt(e);
    };

    window.addEventListener('beforeinstallprompt', handleBeforeInstallPrompt);

    return () => {
      window.removeEventListener('beforeinstallprompt', handleBeforeInstallPrompt);
    };
  }, []);

  const handleInstallClick = async () => {
    if (!deferredPrompt) return;
    deferredPrompt.prompt();
    const { outcome } = await deferredPrompt.userChoice;
    if (outcome === 'accepted') {
      setDeferredPrompt(null);
    }
  };

  // WebSocket Connection for real-time sync
  useEffect(() => {
    const protocol = window.location.protocol === 'https:' ? 'wss:' : 'ws:';
    const wsUrl = `${protocol}//${window.location.host}`;
    const socket = new WebSocket(wsUrl);
    socketRef.current = socket;

    socket.onopen = () => {
      if (spreadsheetId) {
        socket.send(JSON.stringify({ type: 'join', spreadsheetId }));
        // Also sync immediately when connecting/reconnecting with a valid ID
        const savedTokens = localStorage.getItem('google_tokens');
        if (savedTokens) {
          syncFromSheets(JSON.parse(savedTokens), spreadsheetId);
        }
      }
    };

    socket.onmessage = (event) => {
      const data = JSON.parse(event.data);
      if (data.type === 'sync') {
        const savedTokens = localStorage.getItem('google_tokens');
        if (savedTokens && spreadsheetId) {
          syncFromSheets(JSON.parse(savedTokens), spreadsheetId);
        }
      }
    };

    return () => {
      socket.close();
    };
  }, [spreadsheetId]);

  // Load from localStorage on mount
  useEffect(() => {
    const savedInventory = localStorage.getItem('gargam_inventory');
    if (savedInventory) setInventory(JSON.parse(savedInventory));
    
    const savedTokens = localStorage.getItem('google_tokens');
    if (savedTokens) {
      const parsedTokens = JSON.parse(savedTokens);
      setTokens(parsedTokens);
      // Initial sync if tokens exist
      if (spreadsheetId) {
        syncFromSheets(parsedTokens, spreadsheetId);
      }
    }

    const handleOAuthMessage = (event: MessageEvent) => {
      if (event.data?.type === 'OAUTH_AUTH_SUCCESS') {
        setTokens(event.data.tokens);
        localStorage.setItem('google_tokens', JSON.stringify(event.data.tokens));
        if (spreadsheetId) {
          syncFromSheets(event.data.tokens, spreadsheetId);
        }
      }
    };
    window.addEventListener('message', handleOAuthMessage);
    return () => window.removeEventListener('message', handleOAuthMessage);
  }, []);

  const syncFromSheets = async (authTokens: any, sheetId: string) => {
    if (!authTokens || !sheetId) return;
    setIsSyncing(true);
    try {
      const res = await fetch('/api/sheets/load', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ tokens: authTokens, spreadsheetId: sheetId })
      });
      const result = await res.json();
      if (result.data) {
        setInventory(result.data);
      }
    } catch (error) {
      console.error("Error syncing from Sheets:", error);
    } finally {
      setIsSyncing(false);
    }
  };

  // Save to localStorage whenever inventory changes
  useEffect(() => {
    localStorage.setItem('gargam_inventory', JSON.stringify(inventory));
  }, [inventory]);

  const handleConnectGoogle = async () => {
    const res = await fetch('/api/auth/url');
    const { url } = await res.json();
    window.open(url, 'google_oauth', 'width=600,height=700');
  };

  const processFile = async (file: File) => {
    return new Promise<void>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async () => {
        try {
          const base64 = (reader.result as string).split(',')[1];
          const apiKey = process.env.GEMINI_API_KEY;
          if (!apiKey) throw new Error("GEMINI_API_KEY is not defined");

          const ai = new GoogleGenAI({ apiKey });
          const response = await ai.models.generateContent({
            model: "gemini-3-flash-preview",
            contents: [
              {
                parts: [
                  { text: `Actúa como un sistema OCR de alta precisión especializado en el sector de automoción y recambios. Tu objetivo es digitalizar albaranes y facturas de proveedores de recambios.

INSTRUCCIONES CRÍTICAS:
1. Analiza TODAS las páginas del documento. Un solo PDF puede contener VARIOS albaranes o facturas de DIFERENTES proveedores.
2. Identifica y separa cada documento individualmente por su proveedor y número de factura.
3. Para cada documento encontrado, identifica su tabla de artículos.
4. Ignora textos irrelevantes como condiciones legales o publicidad.
5. Si el documento es una foto o escaneo, corrige errores de lectura comunes (ej. '0' por 'O', '1' por 'I').

CAMPOS A EXTRAER (JSON):
Devuelve un objeto con una propiedad "documentos" que sea un array de objetos, cada uno representando un albarán/factura con:
- proveedor: Nombre de la empresa que emite el documento.
- fecha: Fecha de emisión en formato YYYY-MM-DD.
- numero_factura: El número identificativo del albarán o factura.
- items: Array de objetos con:
    - referencia: Código alfanumérico del fabricante o proveedor (Part Number).
    - descripcion: Nombre o detalle del recambio.
    - cantidad: Número de unidades.
    - pvp: Precio unitario base (PVP) antes de descuentos.
    - descuento: Porcentaje de descuento aplicado.
    - precio_final: Precio unitario neto (PVP tras el descuento).
    - total_item: Importe total de la línea (cantidad * precio_final).

REGLAS DE FORMATO:
- Devuelve ÚNICAMENTE el objeto JSON solicitado.
- No inventes datos. Si un campo no existe, usa null o 0.
- Asegúrate de extraer TODOS los artículos de TODAS las páginas, agrupándolos correctamente por su documento de origen.` },
                  { inlineData: { data: base64, mimeType: file.type } }
                ]
              }
            ],
            config: { responseMimeType: "application/json" }
          });

          const resultText = response.text;
          if (!resultText) throw new Error("No response from Gemini");

          const result: ExtractionResult = JSON.parse(resultText);
          const allNewItems: InvoiceItem[] = [];

          if (result.documentos && Array.isArray(result.documentos)) {
            result.documentos.forEach(doc => {
              const items = doc.items || [];
              items.forEach(item => {
                const pvp = item.pvp || 0;
                let precioFinal = item.precio_final || 0;
                const totalItem = item.total_item || 0;
                const cantidad = item.cantidad || 0;

                if (precioFinal === 0 && totalItem > 0 && cantidad > 0) {
                  precioFinal = totalItem / cantidad;
                }

                let descuento = item.descuento || 0;
                if (descuento === 0 && pvp > 0 && precioFinal > 0 && pvp > precioFinal) {
                  descuento = Number(((pvp - precioFinal) / pvp * 100).toFixed(2));
                }

                allNewItems.push({
                  id: Math.random().toString(36).substr(2, 9),
                  referencia: item.referencia || "S/R",
                  descripcion: item.descripcion || "Sin descripción",
                  cantidad: cantidad,
                  pvp: pvp,
                  descuento: descuento,
                  precio_final: precioFinal,
                  total_item: totalItem || (cantidad * precioFinal),
                  beneficio: (pvp - precioFinal) * cantidad,
                  estado: 'stock',
                  proveedor: doc.proveedor || "Proveedor desconocido",
                  fecha: doc.fecha || new Date().toISOString().split('T')[0],
                  numero_factura: doc.numero_factura || "S/N"
                });
              });
            });
          }

          if (tokens && spreadsheetId && allNewItems.length > 0) {
            await fetch('/api/sheets/save', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ tokens, spreadsheetId, data: allNewItems })
            });
          } else if (allNewItems.length > 0) {
            setInventory(prev => [...allNewItems, ...prev]);
          }
          resolve();
        } catch (error) {
          reject(error);
        }
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  };

  const handleManualSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    const precioFinal = manualItem.pvp * (1 - manualItem.descuento / 100);
    const newItem: InvoiceItem = {
      id: Math.random().toString(36).substr(2, 9),
      fecha: manualItem.fecha,
      proveedor: manualItem.proveedor || "Manual",
      descripcion: manualItem.descripcion || "Sin descripción",
      referencia: manualItem.referencia || "S/R",
      cantidad: manualItem.cantidad,
      pvp: manualItem.pvp,
      descuento: manualItem.descuento,
      precio_final: precioFinal,
      total_item: precioFinal * manualItem.cantidad,
      beneficio: (manualItem.pvp - precioFinal) * manualItem.cantidad,
      estado: 'stock',
      numero_factura: "MANUAL"
    };

    if (tokens && spreadsheetId) {
      await fetch('/api/sheets/save', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ tokens, spreadsheetId, data: [newItem] })
      });
      await syncFromSheets(tokens, spreadsheetId);
    } else {
      setInventory(prev => [newItem, ...prev]);
    }

    setManualItem({
      fecha: new Date().toISOString().split('T')[0],
      proveedor: '',
      descripcion: '',
      referencia: '',
      cantidad: 1,
      pvp: 0,
      descuento: 0
    });
    setShowManualForm(false);
    setActiveTab('inventory');
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const MAX_SIZE = 20 * 1024 * 1024; // 20MB
    const fileArray = Array.from(files);
    
    const oversizedFiles = fileArray.filter(f => f.size > MAX_SIZE);
    if (oversizedFiles.length > 0) {
      alert(`Algunos archivos exceden el límite de 20MB: ${oversizedFiles.map(f => f.name).join(', ')}`);
      return;
    }

    setIsUploading(true);
    try {
      const fileArray = Array.from(files);
      await Promise.all(fileArray.map(file => processFile(file)));
      
      if (tokens && spreadsheetId) {
        await syncFromSheets(tokens, spreadsheetId);
      }
      
      setActiveTab('inventory');
    } catch (error) {
      console.error("Error uploading files:", error);
    } finally {
      setIsUploading(false);
    }
  };

  const updateItemStatus = async (id: string, newStatus: ItemStatus) => {
    const updatedInventory = inventory.map(item => item.id === id ? { ...item, estado: newStatus } : item);
    setInventory(updatedInventory);
    
    if (tokens && spreadsheetId) {
      await fetch('/api/sheets/update-all', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ tokens, spreadsheetId, data: updatedInventory })
      });
    }
  };

  const deleteItem = async (id: string) => {
    const updatedInventory = inventory.filter(item => item.id !== id);
    setInventory(updatedInventory);
    
    if (tokens && spreadsheetId) {
      await fetch('/api/sheets/update-all', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ tokens, spreadsheetId, data: updatedInventory })
      });
    }
  };

  const toggleAbonoRecibido = async (id: string) => {
    const updatedInventory = inventory.map(item => 
      item.id === id ? { ...item, abonoRecibido: !item.abonoRecibido } : item
    );
    setInventory(updatedInventory);
    
    if (tokens && spreadsheetId) {
      await fetch('/api/sheets/update-all', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ tokens, spreadsheetId, data: updatedInventory })
      });
    }
  };

  const exportToExcel = () => {
    const dataToExport = filteredInventory.map(item => ({
      Fecha: item.fecha,
      Proveedor: item.proveedor,
      Descripción: item.descripcion,
      Referencia: item.referencia,
      Cantidad: item.cantidad,
      PVP: item.pvp,
      Descuento: `${item.descuento}%`,
      'Precio Neto (Unidad)': item.precio_final,
      'Beneficio por Unidad': item.pvp - item.precio_final,
      Estado: item.estado.toUpperCase(),
      'Abono Recibido': item.estado === 'abono' ? (item.abonoRecibido ? 'SÍ' : 'NO') : '-'
    }));

    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Inventario");
    
    // Generate filename with current date
    const date = new Date().toISOString().split('T')[0];
    XLSX.writeFile(workbook, `Inventario_Gargam_${date}.xlsx`);
  };

  const filteredInventory = inventory.filter(item => {
    const matchesSearch = item.descripcion.toLowerCase().includes(searchTerm.toLowerCase()) || 
                         item.proveedor.toLowerCase().includes(searchTerm.toLowerCase());
    const matchesStatus = filterStatus === 'all' || item.estado === filterStatus;
    return matchesSearch && matchesStatus;
  });

  // Dashboard Calculations
  const totalStockValue = inventory.reduce((acc, item) => acc + (item.estado === 'stock' ? (item.pvp * item.cantidad) : 0), 0);
  const totalProfit = inventory.filter(item => item.estado === 'usado').reduce((acc, item) => acc + item.beneficio, 0);
  const totalItems = inventory.length;
  const usedItemsCount = inventory.filter(item => item.estado === 'usado').length;
  const abonoItemsCount = inventory.filter(item => item.estado === 'abono').length;

  const chartData = inventory
    .filter(item => item.estado === 'stock' || item.estado === 'usado')
    .reduce((acc: any[], item) => {
      const month = new Date(item.fecha).toLocaleString('es-ES', { month: 'short' });
      const existing = acc.find(d => d.name === month);
      if (existing) {
        existing.value += item.total_item;
      } else {
        acc.push({ name: month, value: item.total_item });
      }
      return acc;
    }, []).slice(-6);

  const profitChartData = inventory
    .filter(item => item.estado === 'usado')
    .reduce((acc: any[], item) => {
      const month = new Date(item.fecha).toLocaleString('es-ES', { month: 'short' });
      const existing = acc.find(d => d.name === month);
      if (existing) {
        existing.value += item.beneficio;
      } else {
        acc.push({ name: month, value: item.beneficio });
      }
      return acc;
    }, [])
    .slice(-6);

  const statusData = [
    { name: 'Stock', value: inventory.filter(item => item.estado === 'stock').length, color: '#10b981' },
    { name: 'Usado', value: inventory.filter(item => item.estado === 'usado').length, color: '#3b82f6' },
    { name: 'Abono', value: inventory.filter(item => item.estado === 'abono').length, color: '#f43f5e' },
  ].filter(d => d.value > 0);

  const supplierData = inventory.reduce((acc: any[], item) => {
    const existing = acc.find(d => d.name === item.proveedor);
    if (existing) {
      existing.value += item.total_item;
    } else {
      acc.push({ name: item.proveedor, value: item.total_item });
    }
    return acc;
  }, []).sort((a, b) => b.value - a.value).slice(0, 5); // Top 5 suppliers

  const COLORS = ['#F27D26', '#3b82f6', '#10b981', '#f43f5e', '#8b5cf6', '#ec4899'];

  return (
    <div className="min-h-screen bg-[#0a0a0a] text-gray-200 font-sans selection:bg-[#F27D26]/30">
      {/* Mobile Header */}
      <header className="lg:hidden fixed top-0 left-0 w-full bg-[#111214] border-b border-white/5 z-[60] px-4 py-2 flex items-center justify-between">
        <div className="scale-75 origin-left">
          <Logo />
        </div>
        <button 
          onClick={() => setIsMobileMenuOpen(!isMobileMenuOpen)}
          className="p-2 text-gray-400 hover:text-white transition-colors"
        >
          {isMobileMenuOpen ? <X className="w-6 h-6" /> : <Menu className="w-6 h-6" />}
        </button>
      </header>

      {/* Mobile Menu Dropdown */}
      <AnimatePresence>
        {isMobileMenuOpen && (
          <motion.div
            initial={{ opacity: 0, y: -10 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -10 }}
            className="lg:hidden fixed top-[72px] left-0 w-full bg-[#111214] border-b border-white/5 z-[55] shadow-2xl"
          >
            <nav className="p-4 space-y-2">
              {[
                { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
                { id: 'inventory', label: 'Inventario', icon: Package },
                { id: 'upload', label: 'Subir Albarán', icon: Upload },
                { id: 'settings', label: 'Configuración', icon: SettingsIcon },
              ].map((item) => (
                <button
                  key={item.id}
                  onClick={() => {
                    setActiveTab(item.id as any);
                    setIsMobileMenuOpen(false);
                  }}
                  className={cn(
                    "w-full flex items-center px-4 py-3 rounded-xl transition-all duration-200",
                    activeTab === item.id ? "bg-[#F27D26] text-white" : "text-gray-400 hover:bg-white/5"
                  )}
                >
                  <item.icon className="w-5 h-5 mr-3" />
                  <span className="font-medium">{item.label}</span>
                </button>
              ))}
            </nav>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Sidebar (Desktop) */}
      <aside className="fixed left-0 top-0 h-full w-64 bg-[#111214] border-r border-white/5 z-50 hidden lg:flex flex-col">
        <Logo />
        
        <nav className="flex-1 px-4 space-y-2 mt-4">
          <button 
            onClick={() => setActiveTab('dashboard')}
            className={cn(
              "w-full flex items-center px-4 py-3 rounded-xl transition-all duration-200 group",
              activeTab === 'dashboard' ? "bg-[#F27D26] text-white shadow-lg shadow-[#F27D26]/20" : "text-gray-400 hover:bg-white/5 hover:text-white"
            )}
          >
            <LayoutDashboard className="w-5 h-5 mr-3" />
            <span className="font-medium">Dashboard</span>
          </button>
          
          <button 
            onClick={() => setActiveTab('inventory')}
            className={cn(
              "w-full flex items-center px-4 py-3 rounded-xl transition-all duration-200 group",
              activeTab === 'inventory' ? "bg-[#F27D26] text-white shadow-lg shadow-[#F27D26]/20" : "text-gray-400 hover:bg-white/5 hover:text-white"
            )}
          >
            <Package className="w-5 h-5 mr-3" />
            <span className="font-medium">Inventario</span>
          </button>
          
          <button 
            onClick={() => setActiveTab('upload')}
            className={cn(
              "w-full flex items-center px-4 py-3 rounded-xl transition-all duration-200 group",
              activeTab === 'upload' ? "bg-[#F27D26] text-white shadow-lg shadow-[#F27D26]/20" : "text-gray-400 hover:bg-white/5 hover:text-white"
            )}
          >
            <Upload className="w-5 h-5 mr-3" />
            <span className="font-medium">Subir Albarán</span>
          </button>
          
          <button 
            onClick={() => setActiveTab('settings')}
            className={cn(
              "w-full flex items-center px-4 py-3 rounded-xl transition-all duration-200 group",
              activeTab === 'settings' ? "bg-[#F27D26] text-white shadow-lg shadow-[#F27D26]/20" : "text-gray-400 hover:bg-white/5 hover:text-white"
            )}
          >
            <SettingsIcon className="w-5 h-5 mr-3" />
            <span className="font-medium">Configuración</span>
          </button>
        </nav>

        <div className="p-4 border-t border-white/5">
          <div className="flex items-center p-3 bg-white/2 rounded-xl border border-white/5">
            <div className="w-10 h-10 rounded-full bg-gradient-to-br from-gray-600 to-gray-800 flex items-center justify-center text-white font-bold">
              GP
            </div>
            <div className="ml-3 overflow-hidden">
              <p className="text-sm font-bold text-white truncate">Gargam Admin</p>
              <p className="text-xs text-gray-500 truncate">Taller Mecánico</p>
            </div>
          </div>
        </div>
      </aside>

      {/* Main Content */}
      <main className="lg:ml-64 p-4 lg:p-8 pt-24 lg:pt-8">
        <header className="flex flex-col md:flex-row md:items-center justify-between mb-8 gap-4">
          <div>
            <h1 className="text-3xl font-bold text-white tracking-tight">
              {activeTab === 'dashboard' && 'Resumen Global'}
              {activeTab === 'inventory' && 'Control de Inventario'}
              {activeTab === 'upload' && 'Procesar Albaranes'}
              {activeTab === 'settings' && 'Configuración de Sistema'}
            </h1>
            <p className="text-gray-500 mt-1">
              {activeTab === 'dashboard' && 'Vista general del rendimiento y stock.'}
              {activeTab === 'inventory' && 'Gestiona el estado de tus recambios.'}
              {activeTab === 'upload' && 'Sube imágenes o PDFs para extraer datos con IA.'}
              {activeTab === 'settings' && 'Conexión con Google Sheets y preferencias.'}
            </p>
          </div>

          <div className="flex items-center gap-3">
            {tokens && (
              <button 
                onClick={() => syncFromSheets(tokens, spreadsheetId)}
                disabled={isSyncing}
                className="flex items-center px-3 py-1.5 bg-white/5 border border-white/10 rounded-full text-gray-400 text-xs font-medium hover:bg-white/10 transition-colors disabled:opacity-50"
              >
                <Loader2 className={cn("w-3 h-3 mr-1.5", isSyncing && "animate-spin")} />
                {isSyncing ? 'Sincronizando...' : 'Sincronizar'}
              </button>
            )}
            <div className="flex items-center px-3 py-1.5 bg-emerald-500/10 border border-emerald-500/20 rounded-full text-emerald-400 text-xs font-medium">
              <div className="w-1.5 h-1.5 bg-emerald-500 rounded-full mr-2 animate-pulse"></div>
              Sistema Online
            </div>
            {tokens && (
              <div className="flex items-center px-3 py-1.5 bg-blue-500/10 border border-blue-500/20 rounded-full text-blue-400 text-xs font-medium">
                <CheckCircle2 className="w-3 h-3 mr-1.5" />
                Google Sheets Conectado
              </div>
            )}
          </div>
        </header>

        <AnimatePresence mode="wait">
          {activeTab === 'dashboard' && (
            <motion.div 
              key="dashboard"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="space-y-6"
            >
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-5 gap-6">
                <StatCard 
                  title="Valor en Stock" 
                  value={`${totalStockValue.toLocaleString('es-ES', { minimumFractionDigits: 2 })}€`} 
                  icon={Package}
                  trend="up"
                  trendValue="12% vs mes anterior"
                />
                <StatCard 
                  title="Beneficio Total" 
                  value={`${totalProfit.toLocaleString('es-ES', { minimumFractionDigits: 2 })}€`} 
                  icon={ArrowUpRight}
                  trend="up"
                  trendValue="Ahorro por descuentos"
                />
                <StatCard 
                  title="Total Artículos" 
                  value={totalItems.toString()} 
                  icon={FileText}
                  onClick={() => {
                    setActiveTab('inventory');
                    setFilterStatus('all');
                  }}
                />
                <StatCard 
                  title="Usados (Mes)" 
                  value={usedItemsCount.toString()} 
                  icon={CheckCircle2}
                  trend="up"
                  trendValue="+5"
                  onClick={() => {
                    setActiveTab('inventory');
                    setFilterStatus('usado');
                  }}
                />
                <StatCard 
                  title="Abonos Pendientes" 
                  value={abonoItemsCount.toString()} 
                  icon={AlertCircle}
                  trend="down"
                  trendValue="-2"
                  onClick={() => {
                    setActiveTab('inventory');
                    setFilterStatus('abono');
                  }}
                />
              </div>

              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                <Card title="Gasto en Recambios (Últimos 6 meses)">
                  <div className="h-[300px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={chartData}>
                        <CartesianGrid strokeDasharray="3 3" stroke="#ffffff10" vertical={false} />
                        <XAxis 
                          dataKey="name" 
                          stroke="#9ca3af" 
                          fontSize={12} 
                          tickLine={false} 
                          axisLine={false} 
                        />
                        <YAxis 
                          stroke="#9ca3af" 
                          fontSize={12} 
                          tickLine={false} 
                          axisLine={false}
                          tickFormatter={(value) => `${value}€`}
                        />
                        <Tooltip 
                          contentStyle={{ backgroundColor: '#111214', border: '1px solid #ffffff10', borderRadius: '8px' }}
                          itemStyle={{ color: '#F27D26' }}
                        />
                        <Bar dataKey="value" fill="#F27D26" radius={[4, 4, 0, 0]} barSize={40} />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </Card>

                <Card title="Beneficio Usados (Últimos 6 meses)">
                  <div className="h-[300px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={profitChartData}>
                        <CartesianGrid strokeDasharray="3 3" stroke="#ffffff10" vertical={false} />
                        <XAxis 
                          dataKey="name" 
                          stroke="#9ca3af" 
                          fontSize={12} 
                          tickLine={false} 
                          axisLine={false} 
                        />
                        <YAxis 
                          stroke="#9ca3af" 
                          fontSize={12} 
                          tickLine={false} 
                          axisLine={false}
                          tickFormatter={(value) => `${value}€`}
                        />
                        <Tooltip 
                          contentStyle={{ backgroundColor: '#111214', border: '1px solid #ffffff10', borderRadius: '8px' }}
                          itemStyle={{ color: '#10b981' }}
                        />
                        <Bar dataKey="value" fill="#10b981" radius={[4, 4, 0, 0]} barSize={40} />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </Card>

                <Card title="Distribución de Stock">
                  <div className="h-[300px] w-full flex flex-row items-center gap-4">
                    <div className="flex flex-col gap-3 min-w-[120px]">
                      {statusData.map((d) => (
                        <div key={d.name} className="flex items-center text-xs text-gray-400">
                          <div className="w-2 h-2 rounded-full mr-2 shrink-0" style={{ backgroundColor: d.color }}></div>
                          <span className="truncate">{d.name}</span>
                        </div>
                      ))}
                    </div>
                    <div className="flex-1 h-full">
                      <ResponsiveContainer width="100%" height="100%">
                        <PieChart>
                          <Pie
                            data={statusData}
                            cx="50%"
                            cy="50%"
                            innerRadius={60}
                            outerRadius={80}
                            paddingAngle={5}
                            dataKey="value"
                          >
                            {statusData.map((entry, index) => (
                              <Cell key={`cell-${index}`} fill={entry.color} />
                            ))}
                          </Pie>
                          <Tooltip 
                            contentStyle={{ backgroundColor: '#111214', border: '1px solid #ffffff10', borderRadius: '8px' }}
                          />
                        </PieChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                </Card>

                <Card title="Compras por Proveedor (%)">
                  <div className="h-[300px] w-full flex flex-row items-center gap-4">
                    <div className="flex flex-col gap-2 min-w-[140px] max-h-full overflow-y-auto pr-2 custom-scrollbar">
                      {supplierData.map((d, index) => (
                        <div key={d.name} className="flex items-center text-[10px] text-gray-400">
                          <div className="w-2 h-2 rounded-full mr-2 shrink-0" style={{ backgroundColor: COLORS[index % COLORS.length] }}></div>
                          <span className="truncate" title={d.name}>{d.name}</span>
                        </div>
                      ))}
                    </div>
                    <div className="flex-1 h-full">
                      <ResponsiveContainer width="100%" height="100%">
                        <PieChart>
                          <Pie
                            data={supplierData}
                            cx="50%"
                            cy="50%"
                            innerRadius={60}
                            outerRadius={80}
                            paddingAngle={5}
                            dataKey="value"
                          >
                            {supplierData.map((entry, index) => (
                              <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                            ))}
                          </Pie>
                          <Tooltip 
                            contentStyle={{ backgroundColor: '#111214', border: '1px solid #ffffff10', borderRadius: '8px' }}
                            formatter={(value: number) => `${value.toLocaleString('es-ES', { minimumFractionDigits: 2 })}€`}
                          />
                        </PieChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                </Card>
              </div>
            </motion.div>
          )}

          {activeTab === 'inventory' && (
            <motion.div 
              key="inventory"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="space-y-6"
            >
              <Card>
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-6">
                  <div className="relative flex-1 max-w-md">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-500" />
                    <input 
                      type="text" 
                      placeholder="Buscar por descripción o proveedor..."
                      className="w-full bg-white/5 border border-white/10 rounded-xl py-2 pl-10 pr-4 text-sm focus:outline-none focus:ring-2 focus:ring-[#F27D26]/50 transition-all"
                      value={searchTerm}
                      onChange={(e) => setSearchTerm(e.target.value)}
                    />
                  </div>
                  
                  <div className="flex items-center gap-2">
                    <button
                      onClick={exportToExcel}
                      className="flex items-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-sm font-bold transition-all shadow-lg shadow-emerald-900/20"
                    >
                      <FileText className="w-4 h-4" />
                      Exportar Excel
                    </button>
                    
                    <div className="flex bg-white/5 p-1 rounded-xl border border-white/10">
                      {(['all', 'stock', 'usado', 'abono'] as const).map((s) => (
                        <button
                          key={s}
                          onClick={() => setFilterStatus(s)}
                          className={cn(
                            "px-3 py-1.5 rounded-lg text-xs font-medium transition-all",
                            filterStatus === s ? "bg-[#F27D26] text-white shadow-md" : "text-gray-400 hover:text-white"
                          )}
                        >
                          {s.charAt(0).toUpperCase() + s.slice(1)}
                        </button>
                      ))}
                    </div>
                  </div>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-left">
                    <thead>
                      <tr className="text-xs font-semibold text-gray-500 uppercase tracking-wider border-b border-white/5">
                        <th className="px-4 py-3">Fecha</th>
                        <th className="px-4 py-3">Proveedor</th>
                        <th className="px-4 py-3">Descripción</th>
                        <th className="px-4 py-3">Referencia</th>
                        <th className="px-4 py-3 text-right">Cant.</th>
                        <th className="px-4 py-3 text-right">PVP</th>
                        <th className="px-4 py-3 text-right">Dto.</th>
                        <th className="px-4 py-3 text-right">Precio Neto</th>
                        <th className="px-4 py-3 text-right text-emerald-400">Beneficio/Ud.</th>
                        <th className="px-4 py-3 text-center">Estado</th>
                        {filterStatus === 'abono' && (
                          <th className="px-4 py-3 text-center">Recibido</th>
                        )}
                        <th className="px-4 py-3 text-right">Acciones</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-white/5">
                      {filteredInventory.map((item) => (
                        <tr key={item.id} className="group hover:bg-white/2 transition-colors">
                          <td className="px-4 py-4 text-sm text-gray-400">{item.fecha}</td>
                          <td className="px-4 py-4 text-sm font-medium text-white">{item.proveedor}</td>
                          <td className="px-4 py-4 text-sm text-gray-300 max-w-xs truncate">{item.descripcion}</td>
                          <td className="px-4 py-4 text-sm font-mono text-[#F27D26]">{item.referencia}</td>
                          <td className="px-4 py-4 text-sm text-right text-gray-400">{item.cantidad}</td>
                          <td className="px-4 py-4 text-sm text-right font-mono text-gray-400">{item.pvp.toFixed(2)}€</td>
                          <td className="px-4 py-4 text-sm text-right font-mono text-rose-400/70">-{item.descuento}%</td>
                          <td className="px-4 py-4 text-sm text-right font-mono text-white font-bold">{item.precio_final.toFixed(2)}€</td>
                          <td className="px-4 py-4 text-sm text-right font-mono text-emerald-500 font-bold">+{(item.pvp - item.precio_final).toFixed(2)}€</td>
                          <td className="px-4 py-4">
                            <div className="flex justify-center">
                              <select 
                                value={item.estado}
                                onChange={(e) => updateItemStatus(item.id, e.target.value as ItemStatus)}
                                className={cn(
                                  "text-xs font-bold px-2 py-1 rounded-full border-none focus:ring-0 cursor-pointer appearance-none text-center min-w-[80px]",
                                  item.estado === 'stock' && "bg-emerald-500/10 text-emerald-400",
                                  item.estado === 'usado' && "bg-blue-500/10 text-blue-400",
                                  item.estado === 'abono' && "bg-rose-500/10 text-rose-400"
                                )}
                              >
                                <option value="stock">STOCK</option>
                                <option value="usado">USADO</option>
                                <option value="abono">ABONO</option>
                              </select>
                            </div>
                          </td>
                          {filterStatus === 'abono' && (
                            <td className="px-4 py-4">
                              <div className="flex justify-center">
                                <button
                                  onClick={() => toggleAbonoRecibido(item.id)}
                                  className={cn(
                                    "w-5 h-5 rounded border transition-all flex items-center justify-center",
                                    item.abonoRecibido 
                                      ? "bg-emerald-500 border-emerald-500 text-white" 
                                      : "border-white/20 hover:border-[#F27D26]/50"
                                  )}
                                >
                                  {item.abonoRecibido && <CheckCircle2 className="w-3.5 h-3.5" />}
                                </button>
                              </div>
                            </td>
                          )}
                          <td className="px-4 py-4 text-right">
                            <button 
                              onClick={() => deleteItem(item.id)}
                              className="p-2 text-gray-500 hover:text-rose-400 transition-colors opacity-0 group-hover:opacity-100"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          </td>
                        </tr>
                      ))}
                      {filteredInventory.length === 0 && (
                        <tr>
                          <td colSpan={12} className="px-4 py-12 text-center text-gray-500">
                            No se encontraron artículos que coincidan con los filtros.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </Card>
            </motion.div>
          )}

          {activeTab === 'upload' && (
            <motion.div 
              key="upload"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="max-w-2xl mx-auto"
            >
              <Card className="text-center py-12">
                <div className="mb-6 flex justify-center">
                  <div className="w-20 h-20 bg-[#F27D26]/10 rounded-full flex items-center justify-center border-2 border-dashed border-[#F27D26]/30">
                    <Upload className="w-10 h-10 text-[#F27D26]" />
                  </div>
                </div>
                <h2 className="text-2xl font-bold text-white mb-2">Subir Albarán de Compra</h2>
                <p className="text-gray-500 mb-8 max-w-md mx-auto">
                  Sube una foto o PDF del albarán. Nuestra IA extraerá automáticamente los artículos, precios y proveedor para tu inventario.
                </p>
                
                <div className="flex flex-col sm:flex-row items-center justify-center gap-4">
                  <label className={cn(
                    "relative cursor-pointer bg-[#F27D26] hover:bg-[#d96a1b] text-white font-bold py-4 px-8 rounded-xl transition-all shadow-lg shadow-[#F27D26]/20 flex items-center gap-3 w-full sm:w-auto justify-center",
                    isUploading && "opacity-50 cursor-not-allowed"
                  )}>
                    {isUploading ? <Loader2 className="w-5 h-5 animate-spin" /> : <Plus className="w-5 h-5" />}
                    {isUploading ? 'Procesando...' : 'Seleccionar Archivos'}
                    <input 
                      type="file" 
                      className="hidden" 
                      accept="image/*,application/pdf,.jpg,.jpeg" 
                      onChange={handleFileUpload}
                      disabled={isUploading}
                      multiple
                    />
                  </label>

                  <label className={cn(
                    "relative cursor-pointer bg-white text-black hover:bg-gray-200 font-bold py-4 px-8 rounded-xl transition-all shadow-lg shadow-white/5 flex items-center gap-3 w-full sm:w-auto justify-center",
                    isUploading && "opacity-50 cursor-not-allowed"
                  )}>
                    {isUploading ? <Loader2 className="w-5 h-5 animate-spin" /> : <Camera className="w-5 h-5" />}
                    {isUploading ? 'Procesando...' : 'Hacer Fotos'}
                    <input 
                      type="file" 
                      className="hidden" 
                      accept="image/*" 
                      capture="environment"
                      onChange={handleFileUpload}
                      disabled={isUploading}
                      multiple
                    />
                  </label>

                  <button
                    onClick={() => setShowManualForm(!showManualForm)}
                    className="relative cursor-pointer bg-white/5 border border-white/10 text-white hover:bg-white/10 font-bold py-4 px-8 rounded-xl transition-all flex items-center gap-3 w-full sm:w-auto justify-center"
                  >
                    <Edit3 className="w-5 h-5" />
                    Entrada Manual
                  </button>
                </div>

                <AnimatePresence>
                  {showManualForm && (
                    <motion.div
                      initial={{ opacity: 0, height: 0 }}
                      animate={{ opacity: 1, height: "auto" }}
                      exit={{ opacity: 0, height: 0 }}
                      className="mt-8 overflow-hidden"
                    >
                      <form onSubmit={handleManualSubmit} className="bg-white/2 border border-white/5 rounded-2xl p-6 space-y-4">
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                          <div className="space-y-2">
                            <label className="text-[10px] font-bold text-gray-500 uppercase">Fecha</label>
                            <input 
                              type="date" 
                              required
                              value={manualItem.fecha}
                              onChange={(e) => setManualItem({...manualItem, fecha: e.target.value})}
                              className="w-full bg-white/5 border border-white/10 rounded-xl py-2 px-4 text-sm focus:outline-none focus:ring-2 focus:ring-[#F27D26]/50"
                            />
                          </div>
                          <div className="space-y-2">
                            <label className="text-[10px] font-bold text-gray-500 uppercase">Proveedor</label>
                            <input 
                              type="text" 
                              required
                              placeholder="Nombre del proveedor"
                              value={manualItem.proveedor}
                              onChange={(e) => setManualItem({...manualItem, proveedor: e.target.value})}
                              className="w-full bg-white/5 border border-white/10 rounded-xl py-2 px-4 text-sm focus:outline-none focus:ring-2 focus:ring-[#F27D26]/50"
                            />
                          </div>
                        </div>

                        <div className="space-y-2">
                          <label className="text-[10px] font-bold text-gray-500 uppercase">Descripción</label>
                          <input 
                            type="text" 
                            required
                            placeholder="Descripción del artículo"
                            value={manualItem.descripcion}
                            onChange={(e) => setManualItem({...manualItem, descripcion: e.target.value})}
                            className="w-full bg-white/5 border border-white/10 rounded-xl py-2 px-4 text-sm focus:outline-none focus:ring-2 focus:ring-[#F27D26]/50"
                          />
                        </div>

                        <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                          <div className="space-y-2">
                            <label className="text-[10px] font-bold text-gray-500 uppercase">Referencia</label>
                            <input 
                              type="text" 
                              placeholder="Ref."
                              value={manualItem.referencia}
                              onChange={(e) => setManualItem({...manualItem, referencia: e.target.value})}
                              className="w-full bg-white/5 border border-white/10 rounded-xl py-2 px-4 text-sm focus:outline-none focus:ring-2 focus:ring-[#F27D26]/50"
                            />
                          </div>
                          <div className="space-y-2">
                            <label className="text-[10px] font-bold text-gray-500 uppercase">Cantidad</label>
                            <input 
                              type="number" 
                              required
                              min="1"
                              value={manualItem.cantidad}
                              onChange={(e) => setManualItem({...manualItem, cantidad: parseInt(e.target.value)})}
                              className="w-full bg-white/5 border border-white/10 rounded-xl py-2 px-4 text-sm focus:outline-none focus:ring-2 focus:ring-[#F27D26]/50"
                            />
                          </div>
                          <div className="space-y-2">
                            <label className="text-[10px] font-bold text-gray-500 uppercase">PVP (€)</label>
                            <input 
                              type="number" 
                              required
                              step="0.01"
                              min="0"
                              value={manualItem.pvp}
                              onChange={(e) => setManualItem({...manualItem, pvp: parseFloat(e.target.value)})}
                              className="w-full bg-white/5 border border-white/10 rounded-xl py-2 px-4 text-sm focus:outline-none focus:ring-2 focus:ring-[#F27D26]/50"
                            />
                          </div>
                          <div className="space-y-2">
                            <label className="text-[10px] font-bold text-gray-500 uppercase">Descuento (%)</label>
                            <input 
                              type="number" 
                              step="0.01"
                              min="0"
                              max="100"
                              value={manualItem.descuento}
                              onChange={(e) => setManualItem({...manualItem, descuento: parseFloat(e.target.value)})}
                              className="w-full bg-white/5 border border-white/10 rounded-xl py-2 px-4 text-sm focus:outline-none focus:ring-2 focus:ring-[#F27D26]/50"
                            />
                          </div>
                        </div>

                        <div className="pt-4 flex justify-end gap-3">
                          <button 
                            type="button"
                            onClick={() => setShowManualForm(false)}
                            className="px-6 py-2 rounded-xl text-sm font-bold text-gray-400 hover:text-white transition-colors"
                          >
                            Cancelar
                          </button>
                          <button 
                            type="submit"
                            className="px-8 py-2 bg-[#F27D26] hover:bg-[#d96a1b] text-white font-bold rounded-xl transition-all shadow-lg shadow-[#F27D26]/20"
                          >
                            Añadir al Inventario
                          </button>
                        </div>
                      </form>
                    </motion.div>
                  )}
                </AnimatePresence>
                <p className="text-xs text-gray-600 mt-4">Formatos soportados: JPG, JPEG, PNG, PDF (Máx 20MB - Soporta escaneos)</p>

                {isUploading && (
                  <div className="mt-8 space-y-4">
                    <div className="w-full bg-white/5 h-2 rounded-full overflow-hidden">
                      <motion.div 
                        className="h-full bg-[#F27D26]"
                        initial={{ width: "0%" }}
                        animate={{ width: "100%" }}
                        transition={{ duration: 5, ease: "linear" }}
                      />
                    </div>
                    <p className="text-sm text-[#F27D26] font-medium animate-pulse">Analizando estructura del documento...</p>
                  </div>
                )}
              </Card>

              <div className="mt-8 grid grid-cols-1 md:grid-cols-3 gap-4">
                <div className="p-4 bg-white/2 border border-white/5 rounded-xl flex items-start gap-3">
                  <div className="p-2 bg-emerald-500/10 rounded-lg">
                    <CheckCircle2 className="w-4 h-4 text-emerald-400" />
                  </div>
                  <div>
                    <p className="text-xs font-bold text-white uppercase">Extracción Rápida</p>
                    <p className="text-[10px] text-gray-500">Datos listos en segundos.</p>
                  </div>
                </div>
                <div className="p-4 bg-white/2 border border-white/5 rounded-xl flex items-start gap-3">
                  <div className="p-2 bg-blue-500/10 rounded-lg">
                    <Save className="w-4 h-4 text-blue-400" />
                  </div>
                  <div>
                    <p className="text-xs font-bold text-white uppercase">Auto-Guardado</p>
                    <p className="text-[10px] text-gray-500">Sincronizado con Sheets.</p>
                  </div>
                </div>
                <div className="p-4 bg-white/2 border border-white/5 rounded-xl flex items-start gap-3">
                  <div className="p-2 bg-amber-500/10 rounded-lg">
                    <AlertCircle className="w-4 h-4 text-amber-400" />
                  </div>
                  <div>
                    <p className="text-xs font-bold text-white uppercase">Control de Errores</p>
                    <p className="text-[10px] text-gray-500">Verifica antes de confirmar.</p>
                  </div>
                </div>
              </div>
            </motion.div>
          )}

          {activeTab === 'settings' && (
            <motion.div 
              key="settings"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="max-w-2xl mx-auto space-y-6"
            >
              <Card title="Integración con Google">
                <div className="space-y-6">
                  <div className="flex items-center justify-between">
                    <div>
                      <h4 className="text-white font-bold">Google Sheets API</h4>
                      <p className="text-sm text-gray-500">Conecta tu cuenta para sincronizar el inventario.</p>
                    </div>
                    <button 
                      onClick={handleConnectGoogle}
                      className={cn(
                        "px-6 py-2 rounded-xl font-bold transition-all flex items-center gap-2",
                        tokens ? "bg-emerald-500/10 text-emerald-400 border border-emerald-500/20" : "bg-white text-black hover:bg-gray-200"
                      )}
                    >
                      {tokens ? <CheckCircle2 className="w-4 h-4" /> : <ExternalLink className="w-4 h-4" />}
                      {tokens ? 'Conectado' : 'Conectar Google'}
                    </button>
                  </div>

                  <div className="space-y-2">
                    <label className="text-xs font-bold text-gray-500 uppercase">ID de Hoja de Cálculo</label>
                    <input 
                      type="text" 
                      value={spreadsheetId}
                      onChange={(e) => setSpreadsheetId(e.target.value)}
                      placeholder="Introduce el ID de tu Google Sheet..."
                      className="w-full bg-white/5 border border-white/10 rounded-xl py-3 px-4 text-sm focus:outline-none focus:ring-2 focus:ring-[#F27D26]/50"
                    />
                    <p className="text-[10px] text-gray-600">Puedes encontrar el ID en la URL de tu hoja de Google Sheets.</p>
                  </div>

                  {tokens && (
                    <button 
                      onClick={() => {
                        setTokens(null);
                        localStorage.removeItem('google_tokens');
                      }}
                      className="text-xs text-rose-400 hover:underline flex items-center gap-1"
                    >
                      <LogOut className="w-3 h-3" />
                      Desconectar cuenta
                    </button>
                  )}
                </div>
              </Card>

              <Card title="Preferencias de Aplicación">
                <div className="space-y-4">
                  <div className="flex items-center justify-between p-3 bg-white/2 rounded-xl border border-white/5">
                    <div>
                      <p className="text-sm font-bold text-white">Notificaciones de Stock Bajo</p>
                      <p className="text-xs text-gray-500">Avisar cuando queden menos de 2 unidades.</p>
                    </div>
                    <div className="w-12 h-6 bg-[#F27D26] rounded-full relative">
                      <div className="absolute right-1 top-1 w-4 h-4 bg-white rounded-full"></div>
                    </div>
                  </div>
                  <div className="flex items-center justify-between p-3 bg-white/2 rounded-xl border border-white/5">
                    <div>
                      <p className="text-sm font-bold text-white">Modo Oscuro Forzado</p>
                      <p className="text-xs text-gray-500">Optimizado para pantallas de taller.</p>
                    </div>
                    <div className="w-12 h-6 bg-[#F27D26] rounded-full relative">
                      <div className="absolute right-1 top-1 w-4 h-4 bg-white rounded-full"></div>
                    </div>
                  </div>

                  {deferredPrompt && (
                    <div className="mt-6 p-4 bg-[#F27D26]/10 border border-[#F27D26]/20 rounded-xl flex flex-col gap-3">
                      <div>
                        <p className="text-sm font-bold text-[#F27D26]">Instalar Aplicación</p>
                        <p className="text-xs text-gray-400">Instala la app en tu PC o móvil para acceder más rápido y usarla como una app nativa.</p>
                      </div>
                      <button 
                        onClick={handleInstallClick}
                        className="bg-[#F27D26] text-white text-xs font-bold py-2 px-4 rounded-lg hover:bg-[#F27D26]/80 transition-colors flex items-center justify-center gap-2"
                      >
                        <Download className="w-3.5 h-3.5" />
                        DESCARGAR E INSTALAR
                      </button>
                    </div>
                  )}
                </div>
              </Card>
            </motion.div>
          )}
        </AnimatePresence>
      </main>
    </div>
  );
}
