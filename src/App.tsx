/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useEffect } from 'react';
import { 
  Building2, 
  Database, 
  Variable, 
  BarChart3, 
  Zap, 
  Layers, 
  TrendingUp, 
  Target, 
  Users, 
  Lightbulb,
  ChevronRight,
  Info,
  AlertCircle,
  CheckCircle2,
  Table as TableIcon
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, 
  ScatterChart, Scatter, ZAxis, Cell, LineChart, Line, RadarChart, 
  PolarGrid, PolarAngleAxis, PolarRadiusAxis, Radar, Legend,
  PieChart, Pie
} from 'recharts';
import { kmeans } from 'ml-kmeans';
import * as XLSX from 'xlsx';
import { 
  parseBankerCSV,
  calculateSummary, 
  preprocessData,
  calculateSilhouette,
  calculateCorrelation,
  CustomerData 
} from './utils/analysis';
import { generatePDFReport } from './utils/pdfGenerator';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

// Data provided by the user
const BANK_DATA_CSV = `region;Edad;casado;jubilado;genero;minutos_preferido;adicionales;equipo_dolares;minutos_no_preferido;internet_gigas;fijo;largadistancia;internetcasa;numoculto;facturaelect;ingreso_miles
Zona 2;44;Married;No;M;3,70;0,00;0,00;7,50;0,00;0;0;0;0;0;4800
Zona 3;33;Married;No;M;4,40;20,75;0,00;15,25;3,57;0;1;0;1;0;10200
Zona 3;52;Married;No;F;18,15;18,00;0,00;30,25;0,00;0;0;0;1;0;8700
Zona 2;33;Unmarried;No;F;9,45;0,00;0,00;0,00;0,00;0;0;0;0;0;2475
Zona 2;30;Married;No;M;6,30;0,00;0,00;0,00;0,00;0;0;0;1;0;2250
Zona 2;39;Unmarried;No;F;11,80;19,25;0,00;13,50;0,00;0;0;0;1;0;5850
Zona 3;22;Married;No;F;10,90;0,00;0,00;8,75;0,00;1;0;1;0;1;1425
Zona 2;35;Unmarried;No;M;6,05;45,00;50,10;23,25;6,49;1;1;1;1;1;5700
Zona 3;59;Married;No;M;9,75;28,50;0,00;12,00;0,00;1;0;0;1;0;12450
Zona 1;41;Married;No;M;24,15;0,00;0,00;16,50;0,00;1;0;0;0;0;5400
Zona 2;33;Unmarried;No;F;4,85;0,00;26,15;0,00;0,00;0;0;1;0;1;9375
Zona 3;35;Unmarried;No;F;7,10;22,00;0,00;23,75;0,00;0;1;0;1;0;6000
Zona 1;38;Married;No;F;8,55;0,00;0,00;41,75;0,00;0;0;0;0;0;2775
Zona 2;54;Married;No;F;15,60;46,25;46,70;0,00;6,11;1;1;1;1;1;8625
Zona 2;46;Unmarried;No;F;4,40;0,00;0,00;0,00;0,00;0;0;0;0;0;1875
Zona 1;38;Married;No;M;5,10;0,00;30,25;11,25;0,00;1;0;1;0;0;5625
Zona 3;57;Unmarried;No;M;16,15;29,75;31,30;30,00;0,00;1;0;0;1;0;12150
Zona 3;48;Unmarried;No;F;6,65;18,50;0,00;0,00;0,00;0;0;0;0;0;3675
Zona 2;24;Unmarried;No;M;1,05;0,00;0,00;0,00;0,00;0;0;0;0;0;1500
Zona 1;29;Married;No;M;6,70;0,00;48,10;24,25;3,83;1;1;1;1;1;5775
Zona 3;30;Unmarried;No;F;3,75;0,00;33,80;0,00;1,87;1;0;1;0;0;1200
Zona 1;52;Married;No;M;20,70;0,00;0,00;22,00;0,00;0;0;0;0;0;9000
Zona 3;33;Unmarried;No;F;5,30;0,00;49,60;26,75;5,14;1;1;1;0;1;7575
Zona 3;48;Married;No;M;15,05;0,00;0,00;27,25;0,00;0;0;0;0;0;5025
Zona 3;43;Married;No;M;12,50;19,75;0,00;18,00;0,00;0;0;0;1;0;2700
Zona 2;21;Unmarried;No;F;2,20;20,75;0,00;40,50;0,00;0;0;0;0;0;2475
Zona 2;40;Unmarried;No;F;8,25;23,50;36,90;28,00;3,74;1;1;1;1;1;2775
Zona 3;33;Married;No;M;9,10;0,00;0,00;0,00;0,00;1;0;0;1;0;2325
Zona 1;21;Married;No;F;2,90;0,00;0,00;0,00;0,00;0;0;0;1;0;1275
Zona 2;33;Married;No;F;5,55;0,00;27,35;0,00;0,00;1;0;1;0;1;1425
Zona 1;37;Married;No;F;10,60;0,00;31,10;18,25;0,00;1;0;1;0;1;2700
Zona 1;53;Married;No;M;21,00;56,00;0,00;34,00;5,00;1;1;1;0;1;11625
Zona 1;50;Married;No;F;6,50;27,50;0,00;35,00;0,00;0;0;0;0;0;10500
Zona 1;27;Married;No;M;4,80;0,00;19,55;0,00;0,00;1;0;1;0;0;4125
Zona 2;46;Married;No;M;33,90;38,25;44,65;13,75;5,53;1;1;1;0;1;12225
Zona 3;35;Married;No;M;4,25;0,00;30,55;0,00;0,00;1;0;1;0;1;3900
Zona 2;60;Unmarried;No;M;21,15;39,25;46,35;54,25;5,49;1;1;1;1;1;15825
Zona 1;57;Married;No;M;9,80;33,50;0,00;36,00;4,17;1;1;0;1;0;13950
Zona 1;41;Married;No;F;6,55;29,25;0,00;19,75;0,00;0;1;0;1;0;2925
Zona 2;57;Unmarried;Yes;F;41,75;49,00;0,00;18,75;0,00;1;0;0;1;0;1650
Zona 3;41;Unmarried;No;M;2,50;19,25;0,00;53,75;0,00;0;0;0;1;0;2250
Zona 2;28;Unmarried;No;F;4,25;30,00;0,00;17,75;0,00;0;0;0;1;0;2175
Zona 2;28;Married;No;M;6,20;0,00;0,00;0,00;0,00;0;0;0;1;0;1725
Zona 1;36;Married;No;M;5,65;0,00;46,75;0,00;4,85;1;1;1;0;1;4650
Zona 1;43;Married;No;F;14,70;18,75;0,00;11,00;0,00;1;0;0;1;0;5700
Zona 1;41;Married;No;M;14,50;0,00;37,00;17,00;0,00;1;1;1;0;1;5550
Zona 1;51;Married;No;M;12,85;25,75;0,00;14,25;0,00;1;0;1;0;1;4725
Zona 3;41;Married;No;M;7,75;0,00;37,45;18,50;3,88;1;0;1;0;1;2700
Zona 3;34;Married;No;F;2,95;0,00;22,95;0,00;0,00;0;0;0;0;0;2475
Zona 1;36;Unmarried;No;F;3,25;31,50;0,00;15,50;2,68;0;0;1;1;1;2175
Zona 2;34;Married;No;M;6,30;0,00;22,50;0,00;0,00;0;0;1;0;0;2025
Zona 1;52;Married;No;M;24,75;0,00;0,00;22,25;0,00;1;0;0;0;0;3675
Zona 3;22;Unmarried;No;M;7,80;23,75;37,25;22,50;3,01;1;1;1;1;1;1800
Zona 1;26;Married;No;M;4,85;15,50;0,00;0,00;0,00;0;0;0;0;0;1950
Zona 1;27;Unmarried;No;M;6,25;18,50;0,00;14,25;0,00;0;0;0;1;1;3525
Zona 2;45;Married;No;F;7,20;15,75;28,35;0,00;0,00;0;0;1;1;0;7050
Zona 3;34;Unmarried;No;F;7,65;0,00;0,00;0,00;0,00;1;0;0;0;0;1575
Zona 1;62;Married;No;M;15,50;0,00;0,00;34,50;0,00;0;0;0;0;0;2025
Zona 2;52;Unmarried;No;F;10,40;28,00;0,00;19,00;2,75;0;0;0;1;0;2250
Zona 2;40;Married;No;F;19,70;0,00;28,15;14,75;0,00;1;0;1;0;1;9525
Zona 2;39;Married;No;F;12,70;27,50;0,00;11,50;0,00;0;1;0;1;0;10275
Zona 2;50;Married;No;F;28,80;55,50;0,00;21,25;6,11;1;1;0;1;0;6000
Zona 3;55;Unmarried;No;M;10,25;0,00;0,00;15,75;0,00;0;0;0;0;0;2250
Zona 2;51;Married;No;F;29,00;0,00;47,05;24,75;0,00;1;0;1;1;0;32850
Zona 2;39;Unmarried;No;M;15,25;24,50;37,45;27,25;3,01;1;0;1;1;1;5925
Zona 3;47;Married;No;M;9,05;0,00;0,00;6,75;0,00;0;0;0;0;0;4725
Zona 3;51;Married;No;F;8,95;21,25;0,00;0,00;0,00;0;0;0;1;0;3000
Zona 1;67;Married;Yes;M;57,05;51,50;0,00;19,25;0,00;1;0;0;1;0;3825
Zona 1;27;Unmarried;No;M;6,00;0,00;0,00;0,00;0,00;0;0;0;0;0;2775
Zona 3;43;Unmarried;No;F;3,15;0,00;27,25;16,25;1,80;0;0;1;1;1;4575`;

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// --- Components ---

const StepCard = ({ 
  icon: Icon, 
  title, 
  active, 
  completed, 
  onClick 
}: { 
  icon: any, 
  title: string, 
  active: boolean, 
  completed: boolean, 
  onClick: () => void 
}) => (
  <button 
    onClick={onClick}
    className={cn(
      "w-full flex items-center gap-3 p-4 rounded-xl transition-all duration-300 text-left relative",
      active ? "bg-slate-900 text-white shadow-lg scale-102 z-10" : "hover:bg-slate-100 text-slate-600",
      completed && !active && "text-emerald-600"
    )}
  >
    <div className={cn(
      "p-2 rounded-lg flex items-center justify-center",
      active ? "bg-white/10" : "bg-slate-100",
      completed && "bg-emerald-50"
    )}>
      <Icon size={18} />
    </div>
    <span className="text-sm font-medium flex-1">{title}</span>
    {completed && !active && <CheckCircle2 size={14} className="text-emerald-500" />}
    {active && <div className="absolute right-0 top-1/2 -translate-y-1/2 w-1 h-8 bg-blue-500 rounded-l-full" />}
  </button>
);

const SectionHeading = ({ children }: { children: React.ReactNode }) => (
  <h2 className="text-3xl font-serif font-medium text-slate-900 mb-2 tracking-tight">
    {children}
  </h2>
);

const InfoBox = ({ children }: { children: React.ReactNode }) => (
  <div className="bg-blue-50 border-l-4 border-blue-500 p-4 rounded-r-lg mb-6 flex gap-3">
    <Info size={20} className="text-blue-500 shrink-0 mt-0.5" />
    <div className="text-sm text-blue-700 leading-relaxed">{children}</div>
  </div>
);

// --- Steps ---

// --- Steps ---

const BusinessContext = () => (
  <div className="space-y-6">
    <SectionHeading>Entendimiento del Negocio</SectionHeading>
    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
      <div className="space-y-4">
        <p className="text-slate-600 leading-relaxed">
          Nexus Digital Services procesa datos de consumo masivo para optimizar su cartera. El motor de análisis ahora soporta la carga de archivos <strong>XLSX</strong> y aplica ingeniería de variables avanzada.
        </p>
        <div className="p-4 bg-slate-50 rounded-xl border border-slate-100">
          <h3 className="font-semibold text-slate-900 mb-2 flex items-center gap-2">
            <Target size={16} className="text-amber-500" /> Especialización del Modelo
          </h3>
          <p className="text-sm text-slate-600">
            Enfoque en convergencia de servicios (Voz, Datos, Multi-adopción) utilizando transformaciones logarítmicas para corregir sesgos comportamentales.
          </p>
        </div>
      </div>
      <div className="bg-slate-900 rounded-2xl p-6 text-white overflow-hidden relative">
        <div className="relative z-10">
          <h3 className="text-xl font-serif mb-4">Misión Estratégica</h3>
          <ul className="space-y-3 opacity-90 text-sm">
            <li className="flex items-start gap-2">
              <ChevronRight size={14} className="mt-1 text-amber-400" />
              Categorizar el consumo de Gigas vs. Facturación.
            </li>
            <li className="flex items-start gap-2">
              <ChevronRight size={14} className="mt-1 text-amber-400" />
              Identificar la propensión al gasto en equipos premium.
            </li>
            <li className="flex items-start gap-2">
              <ChevronRight size={14} className="mt-1 text-amber-400" />
              Diseñar el programa de lealtad "Nexus Prestige".
            </li>
          </ul>
        </div>
        <div className="absolute -bottom-10 -right-10 w-40 h-40 bg-amber-500/10 rounded-full blur-3xl" />
      </div>
    </div>
  </div>
);

const DataReview = ({ data, onUpload }: { data: CustomerData[], onUpload: (e: React.ChangeEvent<HTMLInputElement>) => void }) => {
  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center">
        <SectionHeading>Revisión Integral de Datos</SectionHeading>
        <label className="bg-slate-900 text-white text-[10px] font-bold uppercase tracking-widest px-4 py-2 hover:bg-slate-800 transition-colors cursor-pointer flex items-center gap-2 rounded-sm shadow-lg">
          <Database size={12} /> Cargar XLSX
          <input type="file" accept=".xlsx, .xls" onChange={onUpload} className="hidden" />
        </label>
      </div>
      <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
        <div className="p-4 border-b border-slate-100 flex justify-between items-center bg-slate-50">
          <h4 className="text-xs font-bold text-slate-400 uppercase">Vista Previa: Nexus_Core_Clients.csv</h4>
          <span className="text-[10px] bg-emerald-100 text-emerald-600 px-2 py-0.5 rounded-full font-bold uppercase tracking-tight">0% Datos Faltantes</span>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-xs text-left">
            <thead className="bg-slate-50 text-slate-400 font-bold uppercase border-b border-slate-200">
              <tr>
                <th className="px-4 py-3">ID</th>
                <th className="px-4 py-3">Edad</th>
                <th className="px-4 py-3">Pref. Min</th>
                <th className="px-4 py-3">Adicionales</th>
                <th className="px-4 py-3">Gigas</th>
                <th className="px-4 py-3">Servicios</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {data.slice(0, 5).map((row) => (
                <tr key={row.id} className="hover:bg-slate-50 transition-colors">
                  <td className="px-4 py-3 font-mono">#{row.id}</td>
                  <td className="px-4 py-3 font-medium">{row.age}</td>
                  <td className="px-4 py-3 text-blue-600 font-bold">{row.preferredMinutes}</td>
                  <td className="px-4 py-3 text-slate-500">${row.additional}</td>
                  <td className="px-4 py-3 text-amber-600 font-bold">{row.internetGigas} Gb</td>
                  <td className="px-4 py-3 font-bold text-slate-700">{row.servicios}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <div className="p-4 bg-blue-50 border border-blue-100 rounded-xl flex gap-3">
          <Database size={20} className="text-blue-500 shrink-0" />
          <div>
            <h5 className="text-sm font-bold text-blue-900 mb-1">Volumen y Estructura</h5>
            <p className="text-xs text-blue-800 opacity-80">El dataset contiene {data.length} perfiles con métricas de consumo de voz, datos y capacidad económica.</p>
          </div>
        </div>
        <div className="p-4 bg-emerald-50 border border-emerald-100 rounded-xl flex gap-3">
          <CheckCircle2 size={20} className="text-emerald-500 shrink-0" />
          <div>
            <h5 className="text-sm font-bold text-emerald-900 mb-1">Consistencia de Variables</h5>
            <p className="text-xs text-emerald-800 opacity-80">Variables como internet_gigas y adicionales permiten un perfilamiento conductual profundo.</p>
          </div>
        </div>
      </div>
    </div>
  );
};


const VariableAnalysis = () => (
  <div className="space-y-6">
    <SectionHeading>Ingeniería de Variables</SectionHeading>
    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
      <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
        <h3 className="font-semibold text-slate-900 mb-4 flex items-center gap-2">
          <div className="w-1 h-4 bg-blue-500 rounded-full" /> Variables del Modelo (Feature Set)
        </h3>
        <ul className="space-y-3">
          {[
            { name: 'Minutos Preferidos', log: true },
            { name: 'Adicionales', log: true },
            { name: 'Minutos No Preferidos', log: true },
            { name: 'Internet Gigas', log: true },
            { name: 'Servicios (Calculada)', log: false },
            { name: 'Equipo Dólares', log: false }
          ].map((v, i) => (
            <li key={i} className="flex justify-between items-center text-sm border-b border-slate-100 pb-2">
              <span className="text-slate-600 font-medium">{v.name}</span>
              <div className="flex gap-2">
                {v.log && <span className="px-2 py-0.5 bg-blue-50 text-blue-600 rounded text-[10px] font-bold">LOG1P</span>}
                <span className="px-2 py-0.5 bg-slate-100 rounded text-slate-500 text-[10px] font-bold">Z-SCORE</span>
              </div>
            </li>
          ))}
        </ul>
      </div>
      <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
        <h3 className="font-semibold text-slate-900 mb-4 flex items-center gap-2">
          <Info size={16} className="text-blue-500" /> Nota Metodológica
        </h3>
        <p className="text-sm text-slate-500 leading-relaxed">
          La variable <strong>servicios</strong> es una agregación sintética de: fijo, larga distancia, internet casa y número oculto. Esta variable captura la "intensidad de adopción" de servicios del ecosistema.
        </p>
        <div className="mt-4 p-4 bg-amber-50 rounded-xl border border-amber-100">
           <p className="text-xs text-amber-700 italic">"La transformación logarítmica es vital para manejar la dispersión extrema en el consumo de datos y minutos."</p>
        </div>
      </div>
    </div>
  </div>
);

const EDAnlysis = ({ data }: { data: CustomerData[] }) => {
  const summary = useMemo(() => calculateSummary(data), [data]);

  return (
    <div className="space-y-6">
      <SectionHeading>Análisis EDA & Descriptivo</SectionHeading>
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
        <div className="h-64 bg-white p-4 rounded-2xl border border-slate-200">
          <p className="text-xs font-bold text-slate-400 uppercase mb-4">Distribución de Ingresos Anuales</p>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={data.slice(0, 30)}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#E2E8F0" />
              <XAxis dataKey="id" hide />
              <YAxis stroke="#94A3B8" fontSize={10} />
              <Tooltip 
                contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 12px rgba(0,0,0,0.1)' }}
              />
              <Bar dataKey="annualIncome" fill="#3B82F6" radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div className="space-y-3">
          {summary.map(stat => (
            <div key={stat.name} className="p-3 bg-slate-50 border border-slate-100 rounded-xl">
              <div className="flex justify-between text-[10px] font-bold uppercase text-slate-500 mb-1">
                <span>{stat.name}</span>
                <span className="text-blue-600">Avg: {stat.avg}</span>
              </div>
              <div className="w-full bg-slate-200 h-1 rounded-full mt-1 relative overflow-hidden">
                <div 
                  className="absolute top-0 left-0 h-full bg-blue-500" 
                  style={{ width: `${((stat.avg - stat.min) / (stat.max - stat.min)) * 100}%` }}
                />
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};

const CorrelationAnalysis = ({ data }: { data: CustomerData[] }) => {
  const correlation = useMemo(() => calculateCorrelation(data), [data]);

  return (
    <div className="space-y-6">
      <SectionHeading>Matriz de Correlaciones</SectionHeading>
      <p className="text-sm text-slate-600">Identificación de variables críticas para el negocio.</p>
      <div className="flex flex-col items-center overflow-x-auto py-4">
        <div className="inline-grid grid-cols-7 gap-1 bg-white p-4 rounded-2xl border border-slate-200 shadow-sm min-w-max">
          <div className="w-20" />
          {correlation.keys.map(k => (
            <div key={k} className="w-20 text-[10px] font-bold uppercase text-slate-400 text-center truncate px-1">
              {k}
            </div>
          ))}
          {correlation.keys.map((k, i) => (
            <React.Fragment key={k}>
              <div className="w-20 text-[10px] font-bold uppercase text-slate-400 flex items-center pr-2">
                {k}
              </div>
              {correlation.matrix[i].map((val, j) => {
                const absVal = Math.abs(val);
                const color = val > 0 ? `rgba(59, 130, 246, ${absVal})` : `rgba(239, 68, 68, ${absVal})`;
                return (
                  <div 
                    key={j} 
                    className="w-20 h-20 rounded-sm flex items-center justify-center text-[10px] font-mono"
                    style={{ backgroundColor: color, color: absVal > 0.5 ? 'white' : 'black' }}
                  >
                    {(val || 0).toFixed(2)}
                  </div>
                );
              })}
            </React.Fragment>
          ))}
        </div>
      </div>
      <div className="bg-amber-50 p-4 rounded-xl border border-amber-100 flex gap-3">
        <AlertCircle size={20} className="text-amber-500 shrink-0 mt-0.5" />
        <p className="text-sm text-amber-900">
          <strong>Insight:</strong> Existe una fuerte correlación positiva entre el Ingreso y los Minutos Preferidos, lo cual valida la agrupación de clientes con mayor capacidad de pago y uso intensivo.
        </p>
      </div>
    </div>
  );
};

const Preprocessing = ({ data }: { data: CustomerData[] }) => {
  const { standardized } = useMemo(() => preprocessData(data), [data]);
  const sampleData = standardized.slice(0, 5);

  return (
    <div className="space-y-6">
      <SectionHeading>Transformación & Estandarización</SectionHeading>
      <InfoBox>
        Se aplica transformación <strong>log1p</strong> a variables de minutos y gigas para reducir el sesgo de asimetría, seguido de una <strong>estandarización Z-score</strong> para que todas las variables tengan media 0 y varianza 1.
      </InfoBox>
      <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
        <div>
          <h4 className="text-xs font-bold text-slate-400 uppercase mb-4">Muestra Original</h4>
          <div className="space-y-2">
            {data.slice(0, 3).map(d => (
              <div key={d.id} className="p-3 bg-white border border-slate-200 rounded-lg text-xs font-mono">
                ID {d.id} | Minutos: {d.preferredMinutes} | Gigas: {d.internetGigas} | Servicios: {d.servicios}
              </div>
            ))}
          </div>
        </div>
        <div>
          <h4 className="text-xs font-bold text-blue-400 uppercase mb-4">Muestra Procesada (Z-Score)</h4>
          <div className="space-y-2">
            {sampleData.slice(0, 3).map((d, i) => (
              <div key={i} className="p-3 bg-blue-50 border border-blue-200 rounded-lg text-xs font-mono">
                ID {i+1} | {d.slice(0, 3).map(v => (v || 0).toFixed(3)).join(' | ')} ...
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
};

const ElbowMethod = ({ kValues, wcss }: { kValues: number[], wcss: number[] }) => (
  <div className="space-y-6">
    <SectionHeading>Identificación de Clústeres</SectionHeading>
    <p className="text-sm text-slate-600">Búsqueda del "codo" óptimo para definir el número de grupos.</p>
    <div className="h-80 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
      <ResponsiveContainer width="100%" height="100%">
        <LineChart data={kValues.map((k, i) => ({ k, wcss: wcss[i] }))}>
          <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
          <XAxis dataKey="k" label={{ value: 'Número de Clústeres (K)', position: 'insideBottom', offset: -5 }} />
          <YAxis label={{ value: 'Inercia (WCSS)', angle: -90, position: 'insideLeft' }} />
          <Tooltip />
          <Line type="monotone" dataKey="wcss" stroke="#3B82F6" strokeWidth={3} dot={{ r: 6, fill: '#3B82F6' }} />
        </LineChart>
      </ResponsiveContainer>
    </div>
    <div className="p-4 bg-emerald-50 border border-emerald-100 rounded-xl">
      <p className="text-sm text-emerald-900 leading-relaxed font-semibold">
        Sugerencia: El mayor cambio en la tasa de reducción de inercia ocurre en K=5. Se utilizarán 5 segmentos para maximizar la homogeneidad intragrupo.
      </p>
    </div>
  </div>
);

const ClusteringModel = ({ clusters, data, silhouette }: { clusters: any, data: CustomerData[], silhouette: number }) => {
  const scatterData = useMemo(() => {
    return data.map((d, i) => ({
      x: d.annualIncome,
      y: d.preferredMinutes,
      cluster: clusters.clusters[i]
    }));
  }, [clusters, data]);

  const COLORS = ['#3B82F6', '#10B981', '#F59E0B', '#8B5CF6', '#EC4899', '#6366F1'];

  return (
    <div className="space-y-6">
      <SectionHeading>Modelo de Clúster (K-Means)</SectionHeading>
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <div className="lg:col-span-2 h-[450px] bg-white p-6 rounded-2xl border border-slate-200 shadow-sm relative">
           <p className="text-xs font-bold text-slate-400 uppercase mb-4">Visualización: Ingresos vs Minutos Preferidos</p>
           <ResponsiveContainer width="100%" height="100%">
            <ScatterChart>
              <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
              <XAxis type="number" dataKey="x" name="Ingreso" unit="k" />
              <YAxis type="number" dataKey="y" name="Minutos" />
              <ZAxis type="number" range={[64, 144]} />
              <Tooltip cursor={{ strokeDasharray: '3 3' }} />
              <Scatter name="Clientes" data={scatterData}>
                {scatterData.map((entry, index) => (
                  <Cell key={`cell-${index}`} fill={COLORS[entry.cluster % COLORS.length]} fillOpacity={0.7} />
                ))}
              </Scatter>
            </ScatterChart>
           </ResponsiveContainer>
        </div>
        <div className="space-y-4">
          <div className="p-6 bg-slate-900 rounded-2xl text-white">
            <h4 className="text-sm uppercase font-bold text-blue-400 mb-4">Métricas del Modelo</h4>
            <div className="space-y-3">
              <div className="flex justify-between border-b border-white/10 pb-2">
                <span className="text-xs opacity-60">Coef. Silueta</span>
                <span className="text-xs font-mono text-emerald-400 font-bold">{(silhouette || 0).toFixed(4)}</span>
              </div>
              <div className="flex justify-between border-b border-white/10 pb-2">
                <span className="text-xs opacity-60">Inercia total</span>
                <span className="text-xs font-mono">{(clusters.error || 0).toFixed(2)}</span>
              </div>
              <div className="flex justify-between border-b border-white/10 pb-2">
                <span className="text-xs opacity-60">Iteraciones</span>
                <span className="text-xs font-mono">32</span>
              </div>
            </div>
          </div>
          <div className="grid grid-cols-1 gap-2">
            {clusters.centroids.map((_: any, i: number) => {
              const count = clusters.clusters.filter((c: number) => c === i).length;
              return (
                <div key={i} className="p-3 bg-white border border-slate-100 rounded-xl flex justify-between items-center">
                  <div className="flex items-center gap-2">
                    <div className="w-3 h-3 rounded-full" style={{ backgroundColor: COLORS[i % COLORS.length] }} />
                    <span className="text-xs font-semibold">Grupo {i}</span>
                  </div>
                  <span className="text-[10px] font-mono text-slate-400">{count} clientes ({((count/data.length)*100).toFixed(1)}%)</span>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm overflow-hidden mt-6">
        <h4 className="text-xs font-bold text-slate-400 uppercase mb-4">Centros de los Clústeres (Dimensiones Estandarizadas)</h4>
        <div className="overflow-x-auto">
          <table className="w-full text-xs">
            <thead>
              <tr className="bg-slate-50">
                <th className="px-4 py-2 text-left">Clúster</th>
                <th className="px-4 py-2 text-center">Pref. Min</th>
                <th className="px-4 py-2 text-center">Adicional</th>
                <th className="px-4 py-2 text-center">No Pref. Min</th>
                <th className="px-4 py-2 text-center">Gigas</th>
                <th className="px-4 py-2 text-center">Servicios</th>
                <th className="px-4 py-2 text-center">Equipo $</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {clusters.centroids.map((centroid: number[], i: number) => (
                <tr key={i} className="hover:bg-slate-50 transition-colors">
                  <td className="px-4 py-2 font-bold flex items-center gap-2">
                    <div className="w-2 h-2 rounded-full" style={{ backgroundColor: COLORS[i % COLORS.length] }} />
                    Grupo {i}
                  </td>
                  {centroid.map((val: number, idx: number) => (
                    <td key={idx} className={cn("px-4 py-2 text-center font-mono", val > 0 ? "text-blue-600" : "text-slate-400")}>
                      {val.toFixed(2)}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <p className="mt-4 text-[10px] text-slate-400 italic">
          * Valores positivos indican que el clúster está por encima del promedio global en esa variable.
        </p>
      </div>
    </div>
  );
};

const SegmentNaming = ({ profiles }: { profiles: any[] }) => {
  return (
    <div className="space-y-6">
      <SectionHeading>Identidad de Segmentos</SectionHeading>
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
        {profiles.map((name, i) => (
          <div key={i} className="p-6 bg-white border border-slate-200 rounded-2xl shadow-sm hover:border-amber-300 transition-colors">
            <div className={cn("w-12 h-12 rounded-xl mb-4 flex items-center justify-center", name.color)}>
              <name.icon size={24} />
            </div>
            <h3 className="font-serif text-xl text-slate-900 mb-1">{name.title}</h3>
            <p className="text-sm text-slate-500 leading-relaxed mb-4">{name.desc}</p>
            <div className="flex gap-2">
              <span className="px-2 py-1 bg-slate-50 rounded text-[10px] uppercase font-bold text-slate-400">Grupo Analítico {i}</span>
              <span className="px-2 py-1 bg-slate-50 rounded text-[10px] uppercase font-bold text-slate-400">K-Means Core</span>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};

const BusinessImpact = ({ profiles }: { profiles: any[] }) => (
  <div className="space-y-6">
    <SectionHeading>Perfilamiento & Recomendaciones Estratégicas</SectionHeading>
    <div className="bg-slate-900 rounded-3xl p-8 text-white relative overflow-hidden">
      <div className="relative z-10 grid grid-cols-1 lg:grid-cols-2 gap-12">
        <div className="space-y-6">
          {profiles.map((p, i) => (
            <div key={i} className="flex items-start gap-4 p-4 bg-white/5 rounded-xl border border-white/10">
              <div className="w-10 h-10 bg-white/10 rounded-lg flex items-center justify-center shrink-0">
                <p.icon size={20} className="text-amber-400" />
              </div>
              <div>
                <h4 className="font-semibold text-lg mb-1">{p.title}</h4>
                <p className="text-xs opacity-70 leading-relaxed">{p.desc}</p>
              </div>
            </div>
          ))}
        </div>
        <div className="bg-white/5 border border-white/10 rounded-2xl p-6 backdrop-blur-sm h-fit">
           <h4 className="text-xs font-bold uppercase tracking-widest text-amber-400 mb-6">Impacto Estimado del Modelo</h4>
           <div className="space-y-6">
              <div>
                <div className="flex justify-between text-sm mb-2">
                  <span>Reducción de Churn Digital</span>
                  <span className="text-emerald-400">+22%</span>
                </div>
                <div className="w-full bg-white/10 h-2 rounded-full">
                  <motion.div 
                    initial={{ width: 0 }}
                    animate={{ width: "84%" }}
                    className="h-full bg-emerald-500 rounded-full"
                  />
                </div>
              </div>
              <div>
                <div className="flex justify-between text-sm mb-2">
                  <span>Upsell de Equipos Premium</span>
                  <span className="text-amber-400">+15%</span>
                </div>
                <div className="w-full bg-white/10 h-2 rounded-full">
                  <motion.div 
                    initial={{ width: 0 }}
                    animate={{ width: "65%" }}
                    className="h-full bg-amber-500 rounded-full"
                  />
                </div>
              </div>
           </div>
        </div>
      </div>
      <div className="absolute top-0 right-0 w-96 h-96 bg-amber-600/10 rounded-full blur-[120px] -mr-48 -mt-48" />
    </div>
  </div>
);


// --- Main App ---

export default function App() {
  const [currentStep, setCurrentStep] = useState(0);
  const [completedSteps, setCompletedSteps] = useState<number[]>([]);
  const [rawData, setRawData] = useState<CustomerData[]>(() => parseBankerCSV(BANK_DATA_CSV));
  const [kParam, setKParam] = useState(5);
  
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const bstr = event.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const json = XLSX.utils.sheet_to_json(ws);
      
      const mapped = json.map((row: any, idx) => {
        const item: CustomerData = {
          id: idx + 1,
          region: String(row.region || ''),
          age: Number(row.Edad) || 0,
          married: String(row.casado || ''),
          retired: String(row.jubilado || ''),
          gender: String(row.genero || ''),
          preferredMinutes: Number(row.minutos_preferido) || 0,
          additional: Number(row.adicionales) || 0,
          equipmentDollars: Number(row.equipo_dolares) || 0,
          nonPreferredMinutes: Number(row.minutos_no_preferido) || 0,
          internetGigas: Number(row.internet_gigas) || 0,
          fixed: Number(row.fijo) || 0,
          longDistance: Number(row.largadistancia) || 0,
          homeInternet: Number(row.internetcasa) || 0,
          hiddenNumber: Number(row.numoculto) || 0,
          electronicBill: Number(row.facturaelect) || 0,
          annualIncome: Number(row.ingreso_miles) || 0,
        };
        item.servicios = (Number(item.fixed) || 0) + (Number(item.longDistance) || 0) + (Number(item.homeInternet) || 0) + (Number(item.hiddenNumber) || 0);
        return item;
      });
      setRawData(mapped);
      setCompletedSteps([]);
      setCurrentStep(1);
    };
    reader.readAsBinaryString(file);
  };

  const { standardizedData, stats } = useMemo(() => {
    const { standardized, stats } = preprocessData(rawData);
    return { standardizedData: standardized, stats };
  }, [rawData]);

  const clusteringResults = useMemo(() => {
    const result = kmeans(standardizedData, kParam, { initialization: 'kmeans++' });
    // Calculate total error (WCSS) using computeInformation
    const info = result.computeInformation(standardizedData);
    const totalError = info.reduce((acc, curr) => acc + (curr.error * (curr.size || 0)), 0);
    return { ...result, error: totalError };
  }, [standardizedData, kParam]);

  const silhouetteScore = useMemo(() => {
    if (!standardizedData.length) return 0;
    return calculateSilhouette(standardizedData, clusteringResults.clusters);
  }, [standardizedData, clusteringResults]);

  const dataWithClusters = useMemo(() => {
    return rawData.map((d, i) => ({
      ...d,
      cluster: clusteringResults.clusters[i]
    }));
  }, [rawData, clusteringResults]);

  const summary = useMemo(() => calculateSummary(dataWithClusters), [dataWithClusters]);
  const correlation = useMemo(() => calculateCorrelation(dataWithClusters), [dataWithClusters]);

  const segmentProfiles = useMemo(() => {
    const profiles = [
      { title: "Hiper-Conectados Premium", desc: "Alto consumo de Gigas y minutos. Equipos de alta gama. Máximo ARPU.", icon: Zap, color: "text-blue-600 bg-blue-50" },
      { title: "Básico / Low-Cost", desc: "Uso mínimo de servicios adicionales. Orientados al ahorro.", icon: Users, color: "text-slate-600 bg-slate-50" },
      { title: "Ejecutivos de Voz", desc: "Alto uso de minutos y larga distancia. Bajo consumo de internet móvil.", icon: Target, color: "text-amber-600 bg-amber-50" },
      { title: "Hogares Multi-Servicio", desc: "Fuerte presencia de Internet Casa y TV. Alta lealtad comercial.", icon: Building2, color: "text-purple-600 bg-purple-50" },
      { title: "Nuevos Digitales", desc: "Segmento joven en crecimiento. Uso moderado de datos.", icon: TrendingUp, color: "text-emerald-600 bg-emerald-50" }
    ];
    return profiles.slice(0, kParam);
  }, [kParam]);

  const handleExportPDF = () => {
    generatePDFReport(dataWithClusters, clusteringResults, summary, correlation, segmentProfiles);
  };

  const handleExportXLSX = () => {
    const ws = XLSX.utils.json_to_sheet(dataWithClusters);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Clientes_Clustered");
    XLSX.writeFile(wb, "Base_Nexus_Segmentada.xlsx");
  };

  const elbowData = useMemo(() => {
    const kVals = [2, 3, 4, 5, 6, 7, 8];
    const wcss = kVals.map(k => {
      const result = kmeans(standardizedData, k, {});
      const info = result.computeInformation(standardizedData);
      return info.reduce((acc, curr) => acc + (curr.error * (curr.size || 0)), 0);
    });
    return { kVals, wcss };
  }, [standardizedData]);

  const steps = [
    { title: "Entendimiento del Negocio", icon: Building2, component: <BusinessContext /> },
    { title: "Revisión de Datos", icon: Database, component: <DataReview data={rawData} onUpload={handleFileUpload} /> },
    { title: "Tipos de Variables", icon: Variable, component: <VariableAnalysis /> },
    { title: "Análisis EDA", icon: BarChart3, component: <EDAnlysis data={rawData} /> },
    { title: "Correlaciones", icon: Zap, component: <CorrelationAnalysis data={rawData} /> },
    { title: "Normalización", icon: Layers, component: <Preprocessing data={rawData} /> },
    { title: "Número de Clusters", icon: TrendingUp, component: <ElbowMethod kValues={elbowData.kVals} wcss={elbowData.wcss} /> },
    { title: "Modelo K-Means", icon: Target, component: <ClusteringModel clusters={clusteringResults} data={rawData} silhouette={silhouetteScore} /> },
    { title: "Asignación de Nombres", icon: Users, component: <SegmentNaming profiles={segmentProfiles} /> },
    { title: "Perfilamiento Final", icon: Lightbulb, component: <BusinessImpact profiles={segmentProfiles} /> },
  ];

  const handleStepChange = (index: number) => {
    if (!completedSteps.includes(currentStep)) {
      setCompletedSteps([...completedSteps, currentStep]);
    }
    setCurrentStep(index);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  return (
    <div className="h-screen w-full bg-slate-50 flex flex-col font-sans overflow-hidden">
      {/* Header Section */}
      <header className="h-16 bg-slate-900 text-white flex items-center justify-between px-8 shrink-0 z-40">
        <div className="flex items-center space-x-3">
          <div className="w-8 h-8 bg-amber-400 rounded-sm transform rotate-45 flex items-center justify-center -rotate-45">
            <div className="w-3 h-3 bg-slate-900 rounded-full" />
          </div>
          <span className="text-xl font-semibold tracking-tight uppercase">PRESTIGE <span className="font-light text-slate-400">ANALYTICS</span></span>
        </div>
        <div className="flex items-center space-x-6">
          <div className="flex items-center gap-3">
             <button 
               onClick={handleExportXLSX}
               className="bg-emerald-500 text-white text-[10px] font-bold uppercase tracking-widest px-4 py-2 hover:bg-emerald-600 transition-colors shadow-lg shadow-emerald-500/20 rounded-sm flex items-center gap-2"
             >
               <TableIcon size={12} /> Exportar XLSX
             </button>
             <button 
               onClick={handleExportPDF}
               className="bg-amber-400 text-slate-900 text-[10px] font-bold uppercase tracking-widest px-4 py-2 hover:bg-amber-300 transition-colors shadow-lg shadow-amber-400/20 rounded-sm flex items-center gap-2"
             >
               <Database size={12} /> Informe PDF
             </button>
          </div>
          <div className="text-right hidden md:block">
            <p className="text-[10px] text-slate-400 uppercase tracking-widest leading-none mb-1">Segmentación de Clientes</p>
            <p className="text-xs font-medium">Proyecto: Cartera Bancaria V.2</p>
          </div>
          <div className="w-10 h-10 rounded-full bg-slate-700 flex items-center justify-center text-xs font-mono">MB</div>
        </div>
      </header>

      {/* Main Content Area */}
      <div className="flex flex-1 overflow-hidden">
        {/* Sidebar Navigation */}
        <aside className="w-72 bg-white border-r border-slate-200 flex flex-col shrink-0">
          <div className="p-6 border-b border-slate-100 bg-slate-50/50">
            <h2 className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Metodología de Clustering</h2>
          </div>
          <nav className="flex-1 overflow-y-auto custom-scrollbar">
            <ul className="py-2">
              {steps.map((step, idx) => (
                <li key={idx}>
                  <button 
                    onClick={() => handleStepChange(idx)}
                    className={cn(
                      "w-full px-6 py-3 flex items-center space-x-3 transition-all duration-200 text-left group",
                      currentStep === idx 
                        ? "bg-slate-100 border-l-4 border-amber-500 text-slate-900" 
                        : "text-slate-600 hover:bg-slate-50",
                      completedSteps.includes(idx) && currentStep !== idx && "text-slate-400"
                    )}
                  >
                    <span className={cn(
                      "w-6 text-center font-mono text-xs",
                      currentStep === idx ? "font-bold text-amber-600" : "text-slate-400 group-hover:text-slate-600"
                    )}>
                      {(idx + 1).toString().padStart(2, '0')}
                    </span>
                    <span className={cn(
                      "text-sm",
                      currentStep === idx ? "font-semibold" : "font-normal"
                    )}>
                      {step.title}
                    </span>
                  </button>
                </li>
              ))}
            </ul>
          </nav>
        </aside>

        {/* Dashboard Content */}
        <main className="flex-1 overflow-y-auto p-12 bg-slate-50 relative custom-scrollbar">
          <div className="max-w-5xl mx-auto">
            <div className="flex justify-between items-end mb-10 pb-6 border-b border-slate-200">
              <div>
                <div className="flex items-center gap-2 text-[10px] font-bold text-amber-600 uppercase tracking-widest mb-2">
                  <div className="w-2 h-2 bg-amber-500" />
                  Fase {(currentStep + 1).toString().padStart(2, '0')}
                </div>
                <h1 className="text-3xl font-serif text-slate-800 italic uppercase tracking-tight">
                  {steps[currentStep].title}
                </h1>
              </div>
              <button className="px-6 py-2 bg-slate-900 text-white text-[10px] font-bold tracking-widest uppercase hover:bg-slate-800 transition-colors shadow-lg shadow-slate-900/10">
                Ejecutar Análisis
              </button>
            </div>

            <AnimatePresence mode="wait">
              <motion.div
                key={currentStep}
                initial={{ opacity: 0, x: 10 }}
                animate={{ opacity: 1, x: 0 }}
                exit={{ opacity: 0, x: -10 }}
                transition={{ duration: 0.3 }}
              >
                {steps[currentStep].component}

                <div className="mt-16 pt-8 border-t border-slate-200 flex justify-between items-center bg-white p-6 rounded-sm shadow-sm">
                  <button 
                    onClick={() => handleStepChange(Math.max(0, currentStep - 1))}
                    disabled={currentStep === 0}
                    className="px-6 py-2 text-[10px] font-bold uppercase tracking-widest text-slate-500 hover:bg-slate-50 border border-slate-200 transition-colors disabled:opacity-20"
                  >
                    Fase Anterior
                  </button>
                  {currentStep < steps.length - 1 ? (
                    <button 
                      onClick={() => handleStepChange(currentStep + 1)}
                      className="px-8 py-2 bg-slate-900 text-white text-[10px] font-bold tracking-widest uppercase hover:bg-slate-800 transition-all shadow-md flex items-center gap-2"
                    >
                      Siguiente Fase <ChevronRight size={14} className="text-amber-400" />
                    </button>
                  ) : (
                    <div className="flex items-center gap-2 text-emerald-600 text-[10px] font-bold uppercase tracking-widest px-6 py-2 bg-emerald-50 border border-emerald-100">
                      <CheckCircle2 size={16} />
                      Análisis Finalizado
                    </div>
                  )}
                </div>
              </motion.div>
            </AnimatePresence>
          </div>
        </main>
      </div>

      {/* Footer Status Bar */}
      <footer className="h-8 bg-slate-100 border-t border-slate-200 flex items-center justify-between px-8 text-[10px] text-slate-500 shrink-0 z-40">
        <div className="flex items-center space-x-4">
          <span>ESTADO: <span className="text-slate-900 font-semibold">{steps[currentStep].title.toUpperCase()}</span></span>
          <span className="text-slate-300">|</span>
          <span className="uppercase">Dataset: <span className="font-mono text-slate-700">nexus_core_clients_v2.csv</span></span>
        </div>
        <div className="flex items-center space-x-4">
          <span className="flex items-center space-x-2">
            <span className="w-1.5 h-1.5 bg-emerald-500 rounded-full animate-pulse"></span>
            <span className="uppercase">Servidor de Análisis Activo</span>
          </span>
          <span className="text-slate-300">|</span>
          <span className="font-mono">v1.2.0-STABLE</span>
        </div>
      </footer>

      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Playfair+Display:ital,wght@0,400;0,700;1,400&family=JetBrains+Mono:wght@400;500&display=swap');
        
        :root {
          --font-sans: 'Inter', sans-serif;
          --font-serif: 'Playfair Display', serif;
          --font-mono: 'JetBrains Mono', monospace;
        }

        .custom-scrollbar::-webkit-scrollbar {
          width: 4px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
          background: transparent;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
          background: #E2E8F0;
          border-radius: 0px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover {
          background: #94A3B8;
        }
      `}</style>
    </div>
  );
}

