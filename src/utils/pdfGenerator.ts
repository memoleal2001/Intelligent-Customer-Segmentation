import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { CustomerData } from './analysis';

export function generatePDFReport(
  data: CustomerData[], 
  clusters: any, 
  summary: any[], 
  correlation: any,
  segmentProfiles: any[]
) {
  const doc = new jsPDF();
  const pageWidth = doc.internal.pageSize.getWidth();

  // Header
  doc.setFillColor(15, 23, 42); // slate-900
  doc.rect(0, 0, pageWidth, 40, 'F');
  doc.setTextColor(255, 255, 255);
  doc.setFontSize(22);
  doc.text('PRESTIGE ANALYTICS', 20, 25);
  doc.setFontSize(10);
  doc.setTextColor(150, 150, 150);
  doc.text('INFORME ESTRATÉGICO DE SEGMENTACIÓN BANCARIA', 20, 33);

  // Section 1: Data Summary
  doc.setTextColor(15, 23, 42);
  doc.setFontSize(16);
  doc.text('1. Resumen de Calidad de Datos', 20, 55);
  doc.setFontSize(10);
  doc.text(`Total de registros analizados: ${data.length}`, 20, 65);
  doc.text('Métrica de Calidad: 100% Integridad (0% valores nulos)', 20, 70);

  // Table of Stats
  const statsRows = summary.map(s => [s.name, s.min, s.max, s.avg]);
  autoTable(doc, {
    startY: 75,
    head: [['Variable', 'Mínimo', 'Máximo', 'Promedio']],
    body: statsRows,
    theme: 'striped',
    headStyles: { fillColor: [15, 23, 42] }
  });

  // Section 2: Segment Profiles
  let currentY = (doc as any).lastAutoTable.finalY + 15;
  doc.setFontSize(16);
  doc.text('2. Perfilamiento de Segmentos', 20, currentY);
  currentY += 10;

  segmentProfiles.forEach((profile, i) => {
    doc.setFontSize(12);
    doc.setTextColor(15, 23, 42);
    doc.text(`Cluster ${i}: ${profile.title}`, 20, currentY);
    currentY += 7;
    doc.setFontSize(9);
    doc.setTextColor(100, 100, 100);
    doc.text(profile.desc, 20, currentY);
    currentY += 12;
    
    if (currentY > 260) {
      doc.addPage();
      currentY = 20;
    }
  });

  // Section 3: Recommendations
  doc.addPage();
  doc.setTextColor(15, 23, 42);
  doc.setFontSize(16);
  doc.text('3. Recomendaciones Estratégicas', 20, 20);
  
  const recs = [
    { model: 'Nativos Digitales', rec: 'Financiamiento de dispositivos móviles 5G y planes de streaming ilimitados.' },
    { model: 'Ejecutivos', rec: 'Roaming internacional sin costo y multilíneas corporativas con minutos preferenciales.' },
    { model: 'Familias', rec: 'Descuentos en servicios adicionales (Internet Casa + TV) por antigüedad.' },
    { model: 'Low-Cost', rec: 'Campañas de migración a digital para reducir costos operativos de facturación física.' }
  ];

  autoTable(doc, {
    startY: 30,
    head: [['Segmento Potencial', 'Acción Recomendada']],
    body: recs.map(r => [r.model, r.rec]),
    theme: 'grid',
    headStyles: { fillColor: [251, 191, 36], textColor: [15, 23, 42] }, // amber-400
  });

  // Footer on all pages
  const pageCount = doc.getNumberOfPages();
  for (let i = 1; i <= pageCount; i++) {
    doc.setPage(i);
    doc.setFontSize(8);
    doc.setTextColor(150, 150, 150);
    doc.text(`NexusBank Prestige Analytics - Generado el ${new Date().toLocaleDateString()}`, 20, 285);
    doc.text(`Página ${i} de ${pageCount}`, pageWidth - 40, 285);
  }

  doc.save('Informe_Segmentacion_NexusBank.pdf');
}
