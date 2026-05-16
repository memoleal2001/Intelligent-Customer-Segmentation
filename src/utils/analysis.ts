/**
 * Analysis Utilities for Clustering
 */

export interface CustomerData {
  id: number;
  region: string;
  age: number;
  married: string;
  retired: string;
  gender: string;
  preferredMinutes: number;
  additional: number;
  equipmentDollars: number;
  nonPreferredMinutes: number;
  internetGigas: number;
  fixed: number;
  longDistance: number;
  homeInternet: number;
  hiddenNumber: number;
  electronicBill: number;
  annualIncome: number;
  servicios?: number;
  cluster?: number;
  [key: string]: string | number | undefined;
}

export function parseBankerCSV(csv: string): CustomerData[] {
  const lines = csv.trim().split('\n');
  const headers = lines[0].split(';');
  
  return lines.slice(1).map((line, index) => {
    const values = line.split(';');
    const row: any = { id: index + 1 };
    
    headers.forEach((header, i) => {
      let val: any = values[i];
      
      // Clean numeric values (replace , with . and parse)
      if (typeof val === 'string') {
        const cleaned = val.replace(',', '.');
        if (!isNaN(Number(cleaned)) && cleaned.trim() !== '') {
          val = Number(cleaned);
        }
      }
      
      const mapping: { [key: string]: string } = {
        'region': 'region',
        'Edad': 'age',
        'casado': 'married',
        'jubilado': 'retired',
        'genero': 'gender',
        'minutos_preferido': 'preferredMinutes',
        'adicionales': 'additional',
        'equipo_dolares': 'equipmentDollars',
        'minutos_no_preferido': 'nonPreferredMinutes',
        'internet_gigas': 'internetGigas',
        'fijo': 'fixed',
        'largadistancia': 'longDistance',
        'internetcasa': 'homeInternet',
        'numoculto': 'hiddenNumber',
        'facturaelect': 'electronicBill',
        'ingreso_miles': 'annualIncome'
      };
      
      const key = mapping[header] || header;
      row[key] = val;
    });

    // Step 3: Create 'servicios'
    row.servicios = (Number(row.fixed) || 0) + 
                   (Number(row.longDistance) || 0) + 
                   (Number(row.homeInternet) || 0) + 
                   (Number(row.hiddenNumber) || 0);
    
    return row as CustomerData;
  });
}

const FEATURE_KEYS = [
  'preferredMinutes', 
  'additional', 
  'nonPreferredMinutes', 
  'internetGigas',
  'servicios',
  'equipmentDollars'
];

const LOG_TRANSFORM_KEYS = [
  'preferredMinutes', 
  'additional', 
  'nonPreferredMinutes', 
  'internetGigas'
];

export function preprocessData(data: CustomerData[]) {
  // 1. Log1p transformation
  const transformed = data.map(d => {
    const row: any = { ...d };
    LOG_TRANSFORM_KEYS.forEach(key => {
      row[key] = Math.log1p(Number(d[key]) || 0);
    });
    return row;
  });

  // 2. Standardization (Z-score)
  const stats = FEATURE_KEYS.map(key => {
    const values = transformed.map(d => Number(d[key]));
    const mean = values.reduce((a, b) => a + b, 0) / values.length;
    const std = Math.sqrt(values.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / values.length);
    return { key, mean, std };
  });

  const standardized = transformed.map(d => {
    const row: number[] = [];
    FEATURE_KEYS.forEach(key => {
      const s = stats.find(st => st.key === key)!;
      const val = s.std === 0 ? 0 : (Number(d[key]) - s.mean) / s.std;
      row.push(val);
    });
    return row;
  });

  return { standardized, stats };
}

export function calculateSilhouette(data: number[][], labels: number[]) {
  const n = data.length;
  if (n === 0) return 0;
  
  const clusters = Array.from(new Set(labels));
  if (clusters.length < 2) return 0;

  let totalSilhouette = 0;

  for (let i = 0; i < n; i++) {
    const sameClusterIndices = labels.map((l, idx) => l === labels[i] && idx !== i ? idx : -1).filter(idx => idx !== -1);
    
    // a(i): average distance to same cluster
    let a_i = 0;
    if (sameClusterIndices.length > 0) {
      const distances = sameClusterIndices.map(idx => euclideanDistance(data[i], data[idx]));
      a_i = distances.reduce((a, b) => a + b, 0) / distances.length;
    }

    // b(i): minimum average distance to other clusters
    let b_i = Infinity;
    clusters.forEach(c => {
      if (c === labels[i]) return;
      const otherClusterIndices = labels.map((l, idx) => l === c ? idx : -1).filter(idx => idx !== -1);
      const distances = otherClusterIndices.map(idx => euclideanDistance(data[i], data[idx]));
      const avgDist = distances.reduce((a, b) => a + b, 0) / distances.length;
      if (avgDist < b_i) b_i = avgDist;
    });

    const silhouette_i = (b_i - a_i) / Math.max(a_i, b_i);
    totalSilhouette += silhouette_i;
  }

  return totalSilhouette / n;
}

function euclideanDistance(v1: number[], v2: number[]) {
  return Math.sqrt(v1.reduce((acc, val, i) => acc + Math.pow(val - v2[i], 2), 0));
}

const CLUSTERING_SUMMARY_KEYS = ['age', 'preferredMinutes', 'additional', 'equipmentDollars', 'internetGigas', 'annualIncome', 'servicios'];

export function calculateSummary(data: CustomerData[]) {
  return CLUSTERING_SUMMARY_KEYS.map(key => {
    const values = data.map(d => Number(d[key]) || 0);
    const min = Math.min(...values);
    const max = Math.max(...values);
    const avg = values.reduce((a, b) => a + b, 0) / values.length;
    return { name: key, min, max, avg: parseFloat(avg.toFixed(2)) };
  });
}

export function calculateCorrelation(data: CustomerData[]) {
  const keys = ['age', 'preferredMinutes', 'additional', 'equipmentDollars', 'internetGigas', 'annualIncome', 'servicios'];
  const matrix = keys.map(k1 => {
    return keys.map(k2 => {
      const x = data.map(d => Number(d[k1]) || 0);
      const y = data.map(d => Number(d[k2]) || 0);
      return pearsonCorrelation(x, y);
    });
  });
  return { keys, matrix };
}

function pearsonCorrelation(x: number[], y: number[]) {
  const n = x.length;
  let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0, sumY2 = 0;
  for (let i = 0; i < n; i++) {
    sumX += x[i];
    sumY += y[i];
    sumXY += x[i] * y[i];
    sumX2 += x[i] * x[i];
    sumY2 += y[i] * y[i];
  }
  const num = n * sumXY - sumX * sumY;
  const den = Math.sqrt((n * sumX2 - sumX * sumX) * (n * sumY2 - sumY * sumY));
  if (den === 0) return 0;
  return num / den;
}
