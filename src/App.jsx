import { useState, useCallback, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  Upload,
  FileSpreadsheet,
  RotateCcw,
  Package,
  CheckCircle2,
  AlertTriangle,
  Clock,
  TrendingDown,
  Calendar,
  ChevronDown,
  Copy,
  Check,
  Printer,
  Sparkles,
} from 'lucide-react';

/* ---------- Brand assets (logos embedded as base64) ---------- */

const LOGO_MANDAE = "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAQDAwMDAgQDAwMEBAQFBgoGBgUFBgwICQcKDgwPDg4MDQ0PERYTDxAVEQ0NExoTFRcYGRkZDxIbHRsYHRYYGRj/2wBDAQQEBAYFBgsGBgsYEA0QGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBj/wAARCADwAPADASIAAhEBAxEB/8QAHAABAAEFAQEAAAAAAAAAAAAAAAcBBAUGCAMC/8QAPBAAAQMDAwIEBAQDCAEFAAAAAQACAwQFBgcRIRIxCBNBURQiYXEjQlKBFRYyJDNicpGhsdGSF0OCssH/xAAcAQEAAgMBAQEAAAAAAAAAAAAAAwUCBAYBBwj/xAA2EQACAQMCAwYDBgYDAAAAAAAAAQIDBBEhMRJBYQUTFCJRcUKR8AYVMlKhsSNDYsHS4VOB8f/aAAwDAQACEQMRAD8AglERfo4+OhERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREBXcqi2+LTfLZNMq7P5rcaWx0vQBNP8hqC+RsYEY9eX779uFqJ7bqGlcU6ue7ecPD9ySdOUMcSxkoiIpiMIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiICu5HATY9yru2Wy43m6w2200U9bWzuDIoIGF7nH6AKV840TZpvpFHestvdMzJ6yWMU9nhlBMbPzE/rI9duBv3WlcX9GhUjSm/NLZc/wDzqbNO2qVIOpFaLdkOngodyFRSHplo/lup106bXAKS1RH+03SpG0MY/wAP6z9B++ykubulbU3VrPCRhSozrSUaayzSLVarlfLxBarRRTVtZUO6IoIG9bnn6BdF2PSPBtILHBl+ttXBWXF7TJR43CRJ1kfqH5/r+X33Wet14xjTW4uwHQqw/wA15tMPIqr49gfHCfUdfbYbbnbZo25J9Pqvt+I6SXMZdqldXZ3qTV7SU9tY7zGwO/Lx6AehI9PlC47tDtirdtU4Jxg9kvxy/wAY9XqX9rYQoJylhtbt/hj/AJPojXNWZtU890hrc5v8AxnEKQw/w+xgFj6jqlaxsjx68O3BP7D1XNpAD9t+F0Zq7ZNSsm0sq9SNTro20Ma+IWjGWnoI65A0lzT2IYSf1cc7LnLbY7FXX2dx4ZpY0b0jsttM836v1K/tTPepvOq57vrjl0RRERdAVQREQBERAEREAREQBERAEREAREQBERAEREBUE+ikLTPR3LtTblva4BSWph/tF0qgRDEPXp/WfoP9lh9Om4e7UehfnheLDF1SVLGk7v2YSG8c8nYLpUVGaa0WL4Sxws050qo2fNVOAgdUwt7n0HQOk/4fcn053tntOravgp+X1k9l7LeUuhbdn2cKvnnr/St37vkupaUd8xjTPzMF0Hsf81Zk9hbWX57Q+OD3PV249v6R6kqPbhiL8q8Nlz1TvlXcbxllVdW0UZkeXiNvmAEMYO5O/wD0pDocoZDQVOm3hpx4SNjjIuOT1DdwBty/rI2LjsdiR9h7YnSjWO34TohSYlarNJkeX1FfM6mtzIy8NcX/ACvef/wc/Zc5B3FODr0Ytz4ot5fna1b4vyxfJFtJUpvu5y8uHjC8qem35n1MNhehFmxmwszrXG4Ms9qb88Fo6vx6k9wHgcj/ACDn329d3jqM01ktooMbgGneltM356k9MUlREO/b3++3uSvG6Y5RWGePUvxJ5ALpdpAZLfjETg6MbdmeWONhxuO3PJcV71lvzTWO1tyLPq9uAaZUgEsNAHeS6eL8pO+2+/GxPHPyg+sFe6ncT7+rNPHxY8qfpCPxS6skpUo013cI46c31k+S6HjQZQyDr0y8M+OiWQ/h3DKJ2cN9z1kffk/sPVfDf5J0VuZjp2yai6rVjuTzI2CU/wCpG3/kf8Ppc23IL9m9A7T3w92EY5icB8qrySRhY5/v0uPO5HO/Lvsra23bFNMbg/DdG7U7N9QavdtVeZB5rYnep6uwA9gdvclROL1hh5erWfM+tSXwr+lGedpZ22eNF0jHm+pqureG5K3TWsznVzLGnKakxC1WGOQBkDTI3r+Qezd+B29SSudBsXLoTV3CaHGcDrrzqVmL73qTcjE+no2SbspG+YDJwONunqA4A9h6rnvgnnsuy+z0s2uU8rPJYXL8PTrzeSg7VjittjT1y/8Avr0KIiK/KsIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiA2vTm4YvaNRaG55lbjcbVB1SPpA3r85wYelu3+bbvwp9ob5cvEBJc7jl+Q0+IadWKSMTW2F/ll/fpDj2Pb/oKK/DrQUVz8RdgpbhSRVUP4r/Llb1DcRkg7H6qU49GLfZrxf801hvhs+LPuM1VBaGSkPrHdTuguA+nYDk7+i47tyrb+JcZPFRJYe71b0jH82m/sdB2dCr3XFFZi28rZaJbv06GWtd/yHPaN2A6A2b+VsNpyWVuRysMbpB6lp79RH/y9+n0wdwznTzQe3VFg0yoosgy4t8qsv8462Qk92g/f0HHbck8DSdStf7hkFudieB0f8sYnCPLZT04EctQ33cR2B9h+5K3LBNP63M/BTLQYza4H3qvvTRJUOAYehkg5c79ICrp2StqUat2uGnKSXDnV55zlz9tjajcOtNwoPMknrj9Irl77mWrLZiGmNZDnOrl3Oc59XNbNSWph8xsRPLdmnsB6EjbjgK7uWPX3NqFuofiIvjccxeA+bR42x5jc4+gc3v1Eb8cuP0WOq6zTnw+xyVtxq2ZzqTKN3SSu6oqV+379AH/kdvy+kMx5tkWput+PVWYVvx8c1zgj+GI2hjYZRu1rewGyW9nVuM14NqKX42sPHpTj8K67ircQpfw5LV/D19ZPm+hMdDmt11jvkmm+Bz0+BYPbqQz1Tx+HNJTggH6Dfq7b++5Prk7NkVJbppdNfDPj8dXXO+S4ZVVN3jZtwXdZHIHp6ewPr6ZbolPc9aclyvJbzHi2DBscUz2yiN1XG1jT0NHYN3Hc+3APpH+d6+W+2WE4LovbW2GxRjokuLG9E9T6Et9QPqeT9FDSoRvHGFlHiWE2vhT9Zveb6Es6krdOVw8PVZ5tekV8K6ldW7Jp7gmD3Kw118kyvUe5vifV3N7/ADBS7SNc7kn5CQNvfY+gXP3/AAvp8j5Z3zTyPkkcd3OedyT7lfLj1HtsF3HZ1k7SlwSk5Pdt/wBlyXojm7quq0+JRwvr5vqUREVgagREQBERAEREAREQBERAEREAREQBV3KD3Was2IZTkdDU1lgx25XKnpf7+Wkp3SNi43+YgccArCrVhSjxVHhGcISm8RWTC9vVNt/VbkNJdS3Y9/HG4PezQdHm+b8M7+nbfq6e+23rsrbHtNs7yykfVY3id0uEEZLXSxRHpBHpueN/otX7wtsOXeLC6ol8LWylwvL6FdOc1l0+1CpMrgom1c9I2Toic7pBc5hA3+nK882z3KNQskkvOUXSWqlJ+SEcRQt/RGzsB/z67lYe72a62C8SW2922qt9XH/XT1MZjcP2KzOOadZ1l9M6sxrFLncqdp2M0MR8vf26jxuo5xs4z8ZPGcYy/T3JISryj4eOcb46msDvypQsmuGU4voozT7GCLcXzSy1FyafxiHH+mP9H37+2y1g6b5+GVzzhl86KHf4s/BybQbDqPVxxxyrOgwzLbrj8t9teN3SstkRcJKyCmc+JpHfdwG3qvLl2V3BKrJNJp78+R7S8RRb4E02v0MLLNJNO6WaR0kjju57zuSVkMduwsOYWu+GLzhQ1UVT5e+3X0vB2/2WerdKNSrfYDe63CbzDQtb5jpnUx+RvuR3A+4WoQxS1E7IKeJ0ssh2axg3JPsAtiFahcQag047PDIpU6lKSbWGb3qPqzmGqV++Ivda9lG121NbKckQxe3H53fU8/8AC1O6Y/frG2J95slxtwm/ujV00kXmfbqA3W1s0m1KtkNLeq/CLzT0LZY3ulfTn5BuOSO4H3U6+MgAWbC9ht8kv/1Yqen2hb21ehZWaXBLOz2ws8jfla1atKpcVm8rG/PJyb3Vey2vH9M8/wAqtxr8dxG6V9KP/fihPQ7/ACk8H9l4v08zmO31tfLiF6ZTULnMqpzSSBsBbyQ47cbbhXHjrficO8WV1RoeGq4zwvHsa/TU1TV1sdLRwSTTzOEccMbd3PceAAB3KzOW4ZkuDXptpye1y0NU6JsrWP5BaRvwex27H6hSP4d79QY/mtdXjBrjlF8bTF1sFHF5nkv+o/Jv+v0/db3q1muR3vRCeg1Y0ur6K/tqeq23NkPTDCHEk7v55A46fXg+iqLntetSvY28YJw0zqs680s7Lmb1GypztnVctfZ405N+r5HLyIi6EqQiIgCIiAIiIAiIgCIiAIiICoXXvg9q2UOnGcVj4vMbBLHMWfqAiedv9lyCeSuj/Drn+IYfpjmttyO+U9vqq9u1NFIHEy/hPHGwPqQud+1NGVawlTgsvK290W3YtSNO6UpPGj/Y3DRbxE5nm2ubcavsdEbbXCUQQwQhhpy0EjY9yNh67rz1m19yvTrWMYpiFLbqO2UXRJPE+nB89z/ndv7Dn02KhDQrIrLjGvtovd+r46K3wmXrqJASG7xuA7c+q+9e8jsuU683K94/XxV9vlZF0TxggHZoB7jdVH3Fb/eXD3Pk4PTTOce2cFh951fB54/NxeuuMfsdDeI/G7VmVv02vU9OIaq53CGje9hG/kzR+YWk+u23H3PurjX3VK76MWjHsNwGkpLf5sDnid0QcIo2kANaw8bk7kn/ALWja4aqYpdtI8HgxHIaesvFprKeofFG128RjhI3O4A/q2C2q6ZloTr5hVtdnN8/l+80bOeqURSxO2HWGPILXMJ/fhUtC0q06VvO5pSlSi5ZWG8a6PHoWFSvCc6qozSnJRw8/PU2LTrU64an+GbLLleKWGK50dJUU1RLAzpZN+CSHbe+3dYXw53puN+EC+X91Myo+CqamoET+zyGtIB/dfdv1C0HxHRrIMGxLI4o2tpJomSTh7n1kzoj82/TzzsPQeyj3TLULDLL4P8AJMVumQU1Neak1Pk0Tw7qf1MaB2G3OywdlKrTqqnRlGMqkcLD2+vkZK5jCUHKaclGWueZu/h916zHUXUuuxrKW0c9NLTPqIDDCI/J2I+TjuNj68r30f08xqg8UOol2bRxuZZqpooIiARAZQZHED3HYewUGeGvK8ew3WY3jJrpFbqL4GWPzpQSOo7bDgFbvYNdbHh3isyy8io+Oxa+ysElTA0ks6WDpkAPJA+YEf8ASsb/ALKq069xSsoYi4LbRN5WV74yalreU506U7iWWpPffb9snxcPFhnEOrU7RS0X8uxVZp3W0wBz3RB2xPX369ufb6KRfEXZaTMdQ9LrLVNc2juFaWSt32PlExlw+h23Cwt1s/hXgy1+o7spFQ/zvixaKafeKSXv/d9PUNzzsSB+ywHiH1cxq8XzCb9gd7pbhV2epNWWBjto3AsLQ4EDjjbhQUaEat1R8FQlBpSTbTWvDp8vUlnVdOhUVxUUstYWc6Z+tDcfEBrNfdJ7laMLwKmo7c1lIJXSup2vDG77NYwHjbjk7LM2rUOq1M8F+UZBcqSGC4x0tTTVRib0tkkawfMPuCFgr/knh/18xy3XPLsidj11pY9nN84QzRj1j6nNLXN35G3K+rhqPolZvDnkOBYZkEUbY6Wamp45WvMlS8t3L99udyTzx2WmrWPdUaUbeSrRkuJ4fr68yd1Zd5Obqru2nhZ6enIzmjmJ3TFfCbTXDC4bYMovEPxRrKx3THu5x6eo8/0s247brZMEtuoFdjV6x/Wa4Y9d6Gri6YX0z2lxB36mvGwHsQVBmj+rWC3bRWXSPU6skt9LG10dLcASPkLi8fON+h7SeDtttsPuvmI+HHGcRu0n/qXdLxcJoS2j+HqvOfC/cEFojABPoeo7bb9ipK9jWdxVp10+KUspqDk8Z0alnTHoY07mmqUJUmsKOGnLC65WNTnPIaGK1ZbdLZTvLoaarlgYSdyWtcQP9gsdsjju8lzidz3Kb8r6hTi4ximcVJ5ehRERSGIREQBERAEREAREQBERAFeU1quVbHHJR0NRUNlnFMx0bCQ6UjcMH1PsrTut8wTKccsttFPkDbkHUdyhutL8FG2QSuja5hik6nDpB3Yeob7bHha11UlTp8dNZZNQjGUsSeDRXsfHK6ORjmPadiCOxXzwpppNX8cYKSnqrM80bIaOKaMUUJJDYXNqOSdz1OIO/fj0WDyDN8Lq6XDY7dZ55hZ5Yn1kdRTtj85obF1R9QkPUCWO52bwe3dadO+ruWJ0X8+hPK3ppZUyMtj+kpsf0qbqjV/FYrzUVNBaOpk01KXyG2RMMkLBUdbSx0snO8sQ4PIj9Ngri36lYP8AwD+LXKBxqw6lZWUEdBATciyGQPa4k/hxuPQC4An6KJ9pXCjxdy/n/okVpTbx3n18yCdj7JsfZTJadVcVgoJqStsTPwqOmgp3uoGy+Z0QBskTgJY9g6Td3US76hXtFqfpzHVMqau01rpha46NwZb4vKc4F/U0M87jvH+Jud9j8vZZPtGss/wWYq1pv+YiJLBjN7yevkpbHb31j42+ZIeoMbG33e9xDWj7lbXJpHe6V9NR3W+47a7rVx+ZTWqrrfxph2bs5rTEOrbjeQbq2xC52OXGb3jF2uItQrpoamCqkY58TjET+FL0buDSHdwDsR2WeybHaC9ZDALDmtjdidLJIaKaoqPLkoYnO6jEWOHW7pJO3B34Pqsbi6rd9wJ8K9cN8vXbflvpuZU6UHDixl+/9v7kYVlDV0Fynt9bTSQ1UDzHLE8cscO4K8Nv8JUtU+p1jt2e5HdoIKiaOtqaNtPO6CN7nQwyDzSersZGg/68q7k1J0/2p5WWGsPxTYaK407qaENjpmiSN7on9W5kc14PIGzh3Kz8bcR3ovl+xh4ek/jIa6D077Hb7KvQ79Lv9FOtPrJg7LeyE4uIWQ1jSyk+BjlBgbMx0bzJ5g2kbFHtt0nc+uxKsqbWKzy1NN/EqOV8bII+o/AQv/HbVmTzNtxv+Ds3uO231WKv7l/yH8/9GfhqX/J+hD89uraSjpauppJYoKppkgke3YSgOLSR77OYR9wvDpd1beWd/spydrPiYuFPC+x1dbZacTPFDUQxfPK6uNQH77nb8Mkbc7E7cjk+9LqThAsNQ992ugqoJ6YfHvtlN8XWtBme5hi6ulsexa0nqJ5HG3CjXaVylmdF/T9vrlzPfC0m9Kn18yA0V7dqyGvv9bXU1MKWGeoklZAO0bS4kN/ZWSu4yclllc9wiIsjwIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgP/9k=";
const LOGO_ENVIO = "data:image/jpeg;base64,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";
const TABLEAU_URL = "https://tableau.linkedstore.com/#/views/RelatriosLogistica/BaseOPS?:iid=1";

/* ---------- helpers ---------- */

const parseDateBR = (val) => {
  if (val == null || val === '') return null;
  if (val instanceof Date) return isNaN(val) ? null : val;
  if (typeof val === 'number') {
    const d = new Date(Math.round((val - 25569) * 86400 * 1000));
    return isNaN(d) ? null : d;
  }
  const s = String(val).trim();
  // DD/MM/YYYY HH:MM:SS
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{1,2}):(\d{1,2})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5], +m[6]);
  // DD/MM/YYYY HH:MM
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{1,2})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5]);
  // DD/MM/YYYY
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  // YYYY-MM-DD (ISO sem hora, ex.: vindo de CSV) — força horário local
  // pra não deslocar o dia por causa do fuso (UTC vs. UTC-3)
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  const d = new Date(s);
  return isNaN(d) ? null : d;
};

const monthKeyOf = (d) =>
  d ? `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}` : null;

const ptMonthName = (key) => {
  const [y, m] = key.split('-');
  const names = [
    'Janeiro','Fevereiro','Março','Abril','Maio','Junho',
    'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro',
  ];
  return `${names[+m - 1]} de ${y}`;
};

const today = new Date();
today.setHours(0, 0, 0, 0);

const formatTodayBR = () => {
  const d = today;
  return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
};

const fmtShortDate = (d) => {
  if (!d) return '—';
  return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}`;
};

const daysBetween = (a, b) => {
  if (!a || !b) return null;
  return Math.floor((a.getTime() - b.getTime()) / 86400000);
};

const fmtPct = (v) => {
  if (!isFinite(v)) return '0,0%';
  return (v * 100).toFixed(1).replace('.', ',') + '%';
};

const fmtNum = (n) => new Intl.NumberFormat('pt-BR').format(n);

const classify = (row) => {
  const status = String(row.status || '').trim();
  const dataPrim = parseDateBR(row.data_primeira_tentativa_entrega);
  const dataPrev = parseDateBR(row.data_previsao_entrega_cliente);
  const dataEnt = parseDateBR(row.data_entrega);
  const statusLower = status.toLowerCase();

  // Regra 0 (prioridade máxima): Extraviada ou Em devolução são falhas reais de
  // entrega — a mercadoria não chegou ao cliente. Contam SEMPRE como problema,
  // independentemente das datas (inclusive de uma eventual 1ª tentativa no prazo).
  if (statusLower.includes('extravi') || statusLower.includes('devolu')) {
    return 'em_atraso';
  }

  // Regra 1 (prioritária): se houve 1ª tentativa de entrega dentro do prazo,
  // o pedido conta como NO PRAZO — independentemente de quando a entrega final
  // ocorreu. Esses pedidos NUNCA entram em fora_prazo nem em em_atraso.
  if (dataPrim != null && dataPrev != null && dataPrim.getTime() <= dataPrev.getTime()) {
    return 'no_prazo';
  }

  // Regra 2: sem 1ª tentativa registrada no prazo
  if (status === 'Entregue') {
    if (dataEnt != null && dataPrev != null) {
      return dataEnt.getTime() <= dataPrev.getTime() ? 'no_prazo' : 'fora_prazo';
    }
    return 'no_prazo';
  }

  // Pedido ainda não entregue (status diferente de Entregue):
  // se a previsão ainda não venceu, conta como NO PRAZO (em trânsito mas dentro do prazo).
  // Só entra em "em atraso" se a previsão já passou e a 1ª tentativa não ocorreu no prazo.
  if (dataPrev != null) {
    return today.getTime() > dataPrev.getTime() ? 'em_atraso' : 'no_prazo';
  }
  return 'no_prazo';
};

const CATEGORIES = [
  { key: 'no_prazo',    label: 'No prazo',               short: 'No prazo',       icon: CheckCircle2,  color: 'emerald', isProblem: false },
  { key: 'fora_prazo',  label: 'Entregue fora do prazo', short: 'Fora do prazo',  icon: Clock,         color: 'amber',   isProblem: true  },
  { key: 'em_atraso',   label: 'Em atraso',              short: 'Em atraso',      icon: AlertTriangle, color: 'red',     isProblem: true  },
];

const PROBLEM_ROW_KEYS = ['fora_prazo', 'em_atraso'];

const COLOR = {
  emerald: { soft: 'bg-emerald-50', text: 'text-emerald-800', border: 'border-emerald-200', solid: 'bg-emerald-600',  hex: '#059669', softHex: '#ecfdf5', borderHex: '#a7f3d0' },
  amber:   { soft: 'bg-amber-50',   text: 'text-amber-800',   border: 'border-amber-200',   solid: 'bg-amber-500',    hex: '#d97706', softHex: '#fffbeb', borderHex: '#fde68a' },
  red:     { soft: 'bg-red-50',     text: 'text-red-800',     border: 'border-red-200',     solid: 'bg-red-600',      hex: '#dc2626', softHex: '#fef2f2', borderHex: '#fecaca' },
  sky:     { soft: 'bg-sky-50',     text: 'text-sky-800',     border: 'border-sky-200',     solid: 'bg-sky-600',      hex: '#0284c7', softHex: '#f0f9ff', borderHex: '#bae6fd' },
  rose:    { soft: 'bg-rose-50',    text: 'text-rose-800',    border: 'border-rose-200',    solid: 'bg-rose-600',     hex: '#e11d48', softHex: '#fff1f2', borderHex: '#fecdd3' },
  stone:   { soft: 'bg-stone-100',  text: 'text-stone-800',   border: 'border-stone-300',   solid: 'bg-stone-700',    hex: '#57534e', softHex: '#f5f5f4', borderHex: '#d6d3d1' },
};

/* ---------- HTML report builder (for PDF export) ---------- */

const escapeHTML = (s) => String(s ?? '')
  .replace(/&/g, '&amp;')
  .replace(/</g, '&lt;')
  .replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;')
  .replace(/'/g, '&#39;');

const buildReportHTML = (analysis, filename, todayStr, clientName) => {
  const renderBar = (data) => {
    const total = data.total;
    return `<div class="bar">${CATEGORIES.map((c) => {
      const v = data.categories[c.key];
      const pct = total > 0 ? (v / total) * 100 : 0;
      if (pct === 0) return '';
      return `<span style="width:${pct.toFixed(3)}%;background:${COLOR[c.color].hex};display:block;height:100%"></span>`;
    }).join('')}</div>`;
  };

  const renderCategoryGrid = (data) => {
    const total = data.total;
    const savedCount = (data.savedRows || []).length;
    return `<div class="cat-grid">${CATEGORIES.map((c) => {
      const v = data.categories[c.key];
      const pct = total > 0 ? v / total : 0;
      const cl = COLOR[c.color];
      const showSaved = c.key === 'no_prazo' && savedCount > 0;
      return `
        <div class="cat-box" style="background:${cl.softHex};border-color:${cl.borderHex}">
          <span class="cat-dot" style="background:${cl.hex}"></span>
          <div class="cat-info">
            <div class="cat-label">${escapeHTML(c.short)}</div>
            <div class="cat-value">
              <span class="cat-num">${fmtNum(v)}</span>
              <span class="cat-pct">${fmtPct(pct)}</span>
            </div>
            ${showSaved ? `<div class="cat-sub">inclui ${fmtNum(savedCount)} com 1ª tent. no prazo</div>` : ''}
          </div>
        </div>`;
    }).join('')}</div>`;
  };

  const renderMonthOffenders = (offenders, scope) => {
    if (!offenders || offenders.length === 0) return '';
    return `
      <div class="off-block">
        <div class="off-title">Top ${offenders.length} ofensoras${scope === 'month' ? ' do mês' : ''}</div>
        <div class="off-grid">
          ${offenders.map((o, i) => {
            const issues = CATEGORIES
              .filter((c) => c.isProblem)
              .map((c) => ({ ...c, count: o.categories[c.key] }))
              .filter((c) => c.count > 0)
              .sort((a, b) => b.count - a.count);
            const headline = issues[0];
            return `
              <div class="off-card">
                <div class="off-head">
                  <span class="off-rank">${String(i + 1).padStart(2, '0')}</span>
                  <span class="off-name">${escapeHTML(o.name)}</span>
                </div>
                <div class="off-stats">
                  <strong>${fmtNum(o.problems)}</strong> ocorrência${o.problems > 1 ? 's' : ''} ·
                  <strong>${fmtPct(o.problemRate)}</strong> dos próprios pedidos
                </div>
                <div class="off-stats-2">
                  <strong>${fmtPct(o.shareOfTotal)}</strong> do ${scope === 'month' ? 'mês' : 'total'} · vol. ${fmtNum(o.total)}
                </div>
                ${headline ? `
                  <div class="off-headline">
                    <span class="dot-sm" style="background:${COLOR[headline.color].hex}"></span>
                    Principal: ${escapeHTML(headline.label.toLowerCase())} (${fmtNum(headline.count)})
                  </div>` : ''}
              </div>`;
          }).join('')}
        </div>
      </div>`;
  };

  const renderTrackingRow = (r, category, accent) => {
    const code = r.codigo_rastreamento || '—';
    const carrier = r._carrier || r.transportadora || '—';
    const dataPrev = parseDateBR(r.data_previsao_entrega_cliente);
    const dataEnt = parseDateBR(r.data_entrega);
    const dataEnvio = parseDateBR(r.data_hora_envio);
    const dataPrim = parseDateBR(r.data_primeira_tentativa_entrega);
    let detail = '';
    let badge = '';
    let badgeStyle = '';
    if (category === 'fora_prazo') {
      const days = daysBetween(dataEnt, dataPrev);
      detail = `prev. ${fmtShortDate(dataPrev)} → entrega ${fmtShortDate(dataEnt)}`;
      if (dataPrim) detail += ` · 1ª tent. ${fmtShortDate(dataPrim)}`;
      if (days != null) {
        badge = `+${days}d`;
        badgeStyle = 'background:#fffbeb;color:#92400e;border:1px solid #fde68a;';
      }
    } else if (category === 'em_atraso') {
      const days = daysBetween(today, dataPrev);
      detail = `envio ${fmtShortDate(dataEnvio)} · prev. ${fmtShortDate(dataPrev)}`;
      if (dataPrim) detail += ` · 1ª tent. ${fmtShortDate(dataPrim)}`;
      if (days != null) {
        badge = `+${days}d`;
        badgeStyle = 'background:#fef2f2;color:#991b1b;border:1px solid #fecaca;';
      }
    }
    return `
      <div class="rs-row">
        <span class="rs-stripe" style="background:${accent}"></span>
        <span class="rs-code">${escapeHTML(code)}</span>
        <span class="rs-carrier">${escapeHTML(carrier)}</span>
        <span class="rs-detail">${escapeHTML(detail)}</span>
        ${badge ? `<span class="rs-badge" style="${badgeStyle}">${escapeHTML(badge)}</span>` : ''}
      </div>`;
  };

  const renderSavedRow = (r) => {
    const code = r.codigo_rastreamento || '—';
    const carrier = r._carrier || r.transportadora || '—';
    const dataPrim = parseDateBR(r.data_primeira_tentativa_entrega);
    const dataPrev = parseDateBR(r.data_previsao_entrega_cliente);
    const dataEnt = parseDateBR(r.data_entrega);
    const gapDays = daysBetween(dataEnt, dataPrim);
    const lateDays = daysBetween(dataEnt, dataPrev);
    const detail = `1ª tent. ${fmtShortDate(dataPrim)} · prev. ${fmtShortDate(dataPrev)} · entrega ${fmtShortDate(dataEnt)}`;
    return `
      <div class="rs-row">
        <span class="rs-stripe" style="background:#059669"></span>
        <span class="rs-code">${escapeHTML(code)}</span>
        <span class="rs-carrier">${escapeHTML(carrier)}</span>
        <span class="rs-detail">${escapeHTML(detail)}</span>
        ${gapDays != null ? `<span class="rs-badge" style="background:#ecfdf5;color:#065f46;border:1px solid #a7f3d0;">gap ${gapDays}d</span>` : ''}
        ${lateDays != null && lateDays > 0 ? `<span class="rs-badge" style="background:#fffbeb;color:#92400e;border:1px solid #fde68a;">+${lateDays}d</span>` : ''}
      </div>`;
  };

  const renderSaved = (data, opts = {}) => {
    if (!data.savedRows || data.savedRows.length === 0) return '';
    const periodWord = opts.consolidated ? 'do período' : 'do mês';
    return `
      <div class="saved-block">
        <div class="saved-title">Entregas tardias com 1ª tentativa no prazo</div>
        <div class="saved-summary">
          <span class="saved-num">${fmtNum(data.savedRows.length)}</span>
          <span class="saved-explain">pedido${data.savedRows.length > 1 ? 's' : ''} cuja 1ª tentativa de entrega ocorreu dentro do prazo, mas a entrega final foi posterior à previsão. Sob a regra de OTD ajustado, contam como <strong>no prazo</strong>.</span>
        </div>
        <div class="rs-cat saved-list">
          <div class="rs-head">
            <span class="cat-dot" style="background:#059669"></span>
            <strong>${fmtNum(data.savedRows.length)} rastreio${data.savedRows.length > 1 ? 's' : ''} ${periodWord}</strong>
          </div>
          <div>
            ${data.savedRows.map((r) => renderSavedRow(r)).join('')}
          </div>
        </div>
      </div>`;
  };

  const renderRastreios = (data) => {
    if (!data.rows) return '';
    const hasAny = PROBLEM_ROW_KEYS.some((k) => data.rows[k].length > 0);
    if (!hasAny) return '';
    return `
      <div class="rs-block">
        <div class="rs-title">Rastreios com ocorrências</div>
        ${PROBLEM_ROW_KEYS.map((key) => {
          const cat = CATEGORIES.find((c) => c.key === key);
          const rows = data.rows[key] || [];
          if (rows.length === 0) return '';
          const cl = COLOR[cat.color];
          return `
            <div class="rs-cat">
              <div class="rs-head">
                <span class="cat-dot" style="background:${cl.hex}"></span>
                <strong>${escapeHTML(cat.label)}</strong>
                <span class="rs-count">${fmtNum(rows.length)} rastreio${rows.length > 1 ? 's' : ''}</span>
              </div>
              <div>
                ${rows.map((r) => renderTrackingRow(r, key, cl.hex)).join('')}
              </div>
            </div>`;
        }).join('')}
      </div>`;
  };

  const renderBigOffenders = () => {
    if (!analysis.offenders || analysis.offenders.length === 0) {
      return `<div class="empty-good">Nenhuma transportadora apresentou ocorrências no período.</div>`;
    }
    return `
      <div class="big-off-grid">
        ${analysis.offenders.map((o, i) => {
          const issues = CATEGORIES
            .filter((c) => c.isProblem)
            .map((c) => ({ ...c, count: o.categories[c.key] }))
            .filter((c) => c.count > 0)
            .sort((a, b) => b.count - a.count);
          const headline = issues[0];
          return `
            <div class="big-off">
              <div class="big-off-head">
                <div class="big-off-rank">${String(i + 1).padStart(2, '0')}</div>
                <div class="big-off-vol">
                  <div class="tiny-label">Volume</div>
                  <div class="big-off-vol-num">${fmtNum(o.total)} pedidos</div>
                  <div class="big-off-vol-sub">${fmtPct(o.volumeShare)} do total</div>
                </div>
              </div>
              <div class="big-off-name">${escapeHTML(o.name)}</div>
              <div class="big-off-why">
                <div class="tiny-label">Por que é ofensora</div>
                Concentrou <strong>${fmtNum(o.problems)}</strong> ocorrência${o.problems > 1 ? 's' : ''} —
                <strong>${fmtPct(o.problemRate)}</strong> dos pedidos da própria transportadora,
                o que representa <strong style="color:#b45309">${fmtPct(o.shareOfTotal)}</strong> do total geral de ${fmtNum(analysis.overall.total)} pedidos.
                ${headline ? `Principal problema: <strong>${escapeHTML(headline.label.toLowerCase())}</strong> (${fmtNum(headline.count)}).` : ''}
              </div>
              <div class="big-off-issues">
                <div class="tiny-label">Detalhamento das ocorrências</div>
                <ul>
                  ${issues.map((c) => `
                    <li>
                      <span class="big-off-icon" style="background:${COLOR[c.color].hex}"></span>
                      <span class="big-off-issue-label">${escapeHTML(c.label)}</span>
                      <span class="big-off-issue-num">${fmtNum(c.count)} <span class="muted">(${fmtPct(c.count / o.total)})</span></span>
                    </li>`).join('')}
                </ul>
              </div>
            </div>`;
        }).join('')}
      </div>
      <p class="fine-print">
        <strong>Critério:</strong> ranking pelo número absoluto de ocorrências (entregas fora do prazo, em atraso, extravios e devoluções).
        A porcentagem em relação ao total considera o universo de ${fmtNum(analysis.overall.total)} pedidos do período.
      </p>`;
  };

  const css = `
    @page { size: A4; margin: 14mm 12mm; }
    * { box-sizing: border-box; margin: 0; padding: 0; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
    html, body { font-family: 'DM Sans', -apple-system, system-ui, sans-serif; color: #1c1c2e; background: #ffffff; line-height: 1.45; }
    strong { font-weight: 600; }

    .client-banner { text-align: center; margin: -4mm 0 12mm; padding: 14px 0 16px; border-bottom: 2px solid #1c1c2e; }
    .client-banner-label { font-size: 9.5px; text-transform: uppercase; letter-spacing: 0.3em; color: rgba(28,28,46,0.55); margin-bottom: 6px; }
    .client-banner-name { font-family: 'Fraunces', Georgia, serif; font-size: 32px; font-weight: 400; color: #1c1c2e; line-height: 1.1; letter-spacing: -0.01em; }

    .report-header { border-bottom: 1px solid rgba(28,28,46,0.15); padding-bottom: 14px; margin-bottom: 22px; display: flex; align-items: center; gap: 12px; }
    .report-header-text { flex: 1; min-width: 0; }
    .brand-mark { width: 36px; height: 36px; background: #1c1c2e; color: #faf7f2; display: flex; align-items: center; justify-content: center; font-size: 18px; }
    .brand-title { font-family: 'Fraunces', Georgia, serif; font-size: 20px; line-height: 1; }
    .brand-meta { font-size: 10px; text-transform: uppercase; letter-spacing: 0.2em; color: rgba(28,28,46,0.5); margin-top: 4px; }

    .report-footer { margin-top: 36px; padding-top: 22px; border-top: 1px solid rgba(28,28,46,0.15); break-inside: avoid; text-align: center; }
    .footer-partners { display: flex; align-items: center; justify-content: center; gap: 18px; margin-bottom: 10px; }
    .footer-partner { width: 64px; height: 64px; border-radius: 6px; overflow: hidden; border: 1px solid rgba(28,28,46,0.12); display: flex; align-items: center; justify-content: center; background: #ffffff; }
    .footer-logo { width: 100%; height: 100%; object-fit: cover; display: block; }
    .footer-tagline { font-size: 9px; text-transform: uppercase; letter-spacing: 0.3em; color: rgba(28,28,46,0.5); }

    .summary { margin-bottom: 28px; }
    .meta-line { display: flex; gap: 8px; align-items: center; font-size: 11px; color: rgba(28,28,46,0.6); margin-bottom: 12px; }
    .meta-line .ff-mono { font-family: 'JetBrains Mono', ui-monospace, monospace; font-size: 11px; }
    .meta-line .sep { color: rgba(28,28,46,0.3); }
    .kpis { display: grid; grid-template-columns: repeat(3, 1fr); gap: 24px; padding-top: 14px; border-top: 1px solid rgba(28,28,46,0.15); }
    .kpi-label { font-size: 10px; text-transform: uppercase; letter-spacing: 0.2em; color: rgba(28,28,46,0.5); }
    .kpi-value { font-family: 'Fraunces', Georgia, serif; font-size: 52px; font-weight: 300; line-height: 1; margin-top: 6px; letter-spacing: -0.02em; }
    .kpi-value-sm { font-family: 'Fraunces', Georgia, serif; font-size: 26px; margin-top: 8px; }
    .kpi-period { font-family: 'Fraunces', Georgia, serif; font-size: 22px; line-height: 1.15; margin-top: 8px; }
    .kpi-sub { font-size: 11px; color: rgba(28,28,46,0.6); margin-top: 4px; }

    h2.section-h { font-family: 'Fraunces', Georgia, serif; font-size: 26px; font-weight: 400; padding-bottom: 12px; border-bottom: 1px solid rgba(28,28,46,0.15); margin-bottom: 20px; break-after: avoid-page; }
    .method-note { font-size: 11px; color: rgba(28,28,46,0.65); margin: -10px 0 18px; max-width: 75ch; line-height: 1.5; }
    .method-note .ff-mono { background: rgba(28,28,46,0.05); padding: 1px 4px; border-radius: 2px; }

    .month-card { border: 1px solid rgba(28,28,46,0.15); border-radius: 2px; overflow: hidden; margin-bottom: 16px; background: #ffffff; }
    .month-card.consolidated { border-color: rgba(28,28,46,0.4); }
    .page-break { break-before: page; page-break-before: always; }

    .month-head { padding: 14px 18px; display: flex; align-items: baseline; justify-content: space-between; border-bottom: 1px solid rgba(28,28,46,0.1); break-inside: avoid; break-after: avoid-page; }
    .month-eyebrow { font-size: 10px; text-transform: uppercase; letter-spacing: 0.25em; color: rgba(28,28,46,0.5); }
    .month-title { font-family: 'Fraunces', Georgia, serif; font-size: 26px; font-weight: 400; margin-top: 4px; }
    .month-total { font-family: 'Fraunces', Georgia, serif; font-size: 32px; font-weight: 300; margin-top: 4px; line-height: 1; text-align: right; }

    .bar { display: flex; height: 6px; margin: 14px 18px 0; border-radius: 999px; overflow: hidden; background: rgba(28,28,46,0.08); }

    .cat-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; padding: 14px 18px 16px; }
    .cat-box { border: 1px solid; border-radius: 2px; padding: 9px 11px; display: flex; align-items: center; gap: 10px; break-inside: avoid; }
    .cat-dot { width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; display: inline-block; }
    .dot-sm { width: 6px; height: 6px; border-radius: 50%; flex-shrink: 0; display: inline-block; }
    .cat-info { flex: 1; min-width: 0; }
    .cat-label { font-size: 10px; text-transform: uppercase; letter-spacing: 0.06em; color: rgba(28,28,46,0.65); }
    .cat-value { display: flex; align-items: baseline; gap: 8px; margin-top: 2px; }
    .cat-num { font-family: 'JetBrains Mono', monospace; font-size: 16px; font-weight: 500; }
    .cat-pct { font-family: 'JetBrains Mono', monospace; font-size: 11px; color: rgba(28,28,46,0.6); }
    .cat-sub { font-size: 9px; color: rgba(28,28,46,0.55); margin-top: 2px; line-height: 1.2; }

    .off-block { border-top: 1px solid rgba(28,28,46,0.1); background: rgba(250,247,242,0.55); padding: 14px 18px; }
    .off-title { font-size: 10px; text-transform: uppercase; letter-spacing: 0.25em; color: rgba(28,28,46,0.6); margin-bottom: 10px; }
    .off-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 6px; }
    .off-card { border: 1px solid rgba(28,28,46,0.15); background: #ffffff; padding: 10px 11px; break-inside: avoid; }
    .off-head { display: flex; align-items: baseline; gap: 8px; margin-bottom: 4px; }
    .off-rank { font-family: 'Fraunces', Georgia, serif; font-size: 22px; font-weight: 300; color: #b45309; line-height: 1; }
    .off-name { font-family: 'Fraunces', Georgia, serif; font-size: 14px; }
    .off-stats { font-size: 11px; color: rgba(28,28,46,0.75); line-height: 1.45; }
    .off-stats-2 { font-size: 11px; color: rgba(28,28,46,0.7); margin-top: 2px; }
    .off-headline { margin-top: 7px; padding-top: 7px; border-top: 1px solid rgba(28,28,46,0.1); display: flex; align-items: center; gap: 6px; font-size: 10px; color: rgba(28,28,46,0.7); }

    .rs-block { border-top: 1px solid rgba(28,28,46,0.1); padding: 14px 18px 16px; }
    .saved-block { border-top: 1px solid rgba(28,28,46,0.1); padding: 14px 18px 16px; background: rgba(236,253,245,0.4); break-inside: avoid; }
    .saved-title { font-size: 10px; text-transform: uppercase; letter-spacing: 0.25em; color: rgba(28,28,46,0.6); margin-bottom: 8px; }
    .saved-summary { display: flex; align-items: baseline; gap: 12px; margin-bottom: 12px; }
    .saved-num { font-family: 'Fraunces', Georgia, serif; font-size: 32px; font-weight: 300; color: #047857; line-height: 1; flex-shrink: 0; }
    .saved-explain { font-size: 11px; line-height: 1.5; color: rgba(28,28,46,0.75); max-width: 60ch; }
    .saved-list { border-color: #a7f3d0; }
    .saved-list .rs-head { border-bottom-color: #d1fae5; }
    .saved-fine { font-size: 9.5px; color: rgba(28,28,46,0.6); margin-top: 10px; line-height: 1.5; max-width: 75ch; }
    .rs-title { font-size: 10px; text-transform: uppercase; letter-spacing: 0.25em; color: rgba(28,28,46,0.6); margin-bottom: 10px; }
    .rs-cat { border: 1px solid rgba(28,28,46,0.15); border-radius: 2px; margin-bottom: 8px; overflow: hidden; background: #ffffff; }
    .rs-head { padding: 8px 12px; display: flex; align-items: center; gap: 10px; border-bottom: 1px solid rgba(28,28,46,0.1); break-after: avoid-page; }
    .rs-head strong { font-size: 12px; font-weight: 500; }
    .rs-count { font-family: 'JetBrains Mono', monospace; font-size: 11px; color: rgba(28,28,46,0.6); margin-left: auto; }
    .rs-row { padding: 6px 12px; display: flex; flex-wrap: wrap; align-items: center; gap: 4px 12px; border-bottom: 1px solid rgba(28,28,46,0.05); break-inside: avoid; }
    .rs-row:last-child { border-bottom: none; }
    .rs-stripe { width: 3px; height: 14px; border-radius: 2px; flex-shrink: 0; display: inline-block; }
    .rs-code { font-family: 'JetBrains Mono', monospace; font-weight: 500; font-size: 11.5px; color: #1c1c2e; }
    .rs-carrier { font-size: 11px; color: rgba(28,28,46,0.65); }
    .rs-detail { font-size: 11px; color: rgba(28,28,46,0.65); flex: 1; min-width: 0; }
    .rs-badge { font-family: 'JetBrains Mono', monospace; font-size: 10px; padding: 1.5px 7px; border-radius: 2px; flex-shrink: 0; }

    .big-off-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; }
    .big-off { border: 1px solid rgba(28,28,46,0.2); background: #ffffff; break-inside: avoid; display: flex; flex-direction: column; }
    .big-off-head { padding: 14px 14px 10px; border-bottom: 1px solid rgba(28,28,46,0.1); display: flex; align-items: flex-start; justify-content: space-between; gap: 10px; }
    .big-off-rank { font-family: 'Fraunces', Georgia, serif; font-size: 38px; font-weight: 300; color: #b45309; line-height: 1; }
    .big-off-vol { text-align: right; }
    .big-off-vol-num { font-family: 'JetBrains Mono', monospace; font-size: 12px; }
    .big-off-vol-sub { font-family: 'JetBrains Mono', monospace; font-size: 10px; color: rgba(28,28,46,0.5); }
    .big-off-name { font-family: 'Fraunces', Georgia, serif; font-size: 17px; padding: 0 14px 12px; }
    .big-off-why { background: rgba(250,247,242,0.6); padding: 10px 14px; font-size: 11px; line-height: 1.5; border-top: 1px solid rgba(28,28,46,0.1); border-bottom: 1px solid rgba(28,28,46,0.1); }
    .big-off-issues { padding: 10px 14px; flex: 1; }
    .big-off-issues ul { list-style: none; }
    .big-off-issues li { display: flex; align-items: center; gap: 8px; padding: 3px 0; font-size: 11px; }
    .big-off-icon { width: 18px; height: 18px; border-radius: 50%; flex-shrink: 0; }
    .big-off-issue-label { flex: 1; }
    .big-off-issue-num { font-family: 'JetBrains Mono', monospace; }
    .muted { color: rgba(28,28,46,0.5); }

    .tiny-label { font-size: 9px; text-transform: uppercase; letter-spacing: 0.2em; color: rgba(28,28,46,0.5); margin-bottom: 3px; }
    .fine-print { font-size: 10.5px; color: rgba(28,28,46,0.6); margin-top: 12px; line-height: 1.5; max-width: 70ch; }
    .empty-good { padding: 18px; border: 1px solid #a7f3d0; background: #ecfdf5; color: #064e3b; border-radius: 2px; font-size: 13px; }
  `;

  const summarySection = `
    <section class="summary">
      <div class="meta-line">
        <span class="ff-mono">${escapeHTML(filename)}</span>
        <span class="sep">·</span>
        <span>Análise em ${escapeHTML(todayStr)}</span>
      </div>
      <div class="kpis">
        <div>
          <div class="kpi-label">Total de encomendas</div>
          <div class="kpi-value">${fmtNum(analysis.overall.total)}</div>
        </div>
        <div>
          <div class="kpi-label">Período coberto</div>
          <div class="kpi-period">${
            analysis.sortedMonths.length === 0
              ? '—'
              : analysis.sortedMonths.length === 1
              ? escapeHTML(ptMonthName(analysis.sortedMonths[0]))
              : `${escapeHTML(ptMonthName(analysis.sortedMonths[0]))} → ${escapeHTML(ptMonthName(analysis.sortedMonths[analysis.sortedMonths.length - 1]))}`
          }</div>
          <div class="kpi-sub">${analysis.sortedMonths.length} ${analysis.sortedMonths.length === 1 ? 'mês' : 'meses'}</div>
        </div>
        <div>
          <div class="kpi-label">Transportadoras</div>
          <div class="kpi-value-sm">${Object.keys(analysis.overall.byCarrier).length}</div>
          <div class="kpi-sub">parceiros operando no período</div>
        </div>
      </div>
    </section>`;

  const monthSections = analysis.sortedMonths.map((mk, idx) => {
    const m = analysis.months[mk];
    return `
      <section class="month-card${idx > 0 ? ' page-break' : ''}">
        <div class="month-head">
          <div>
            <div class="month-eyebrow">Mês</div>
            <div class="month-title">${escapeHTML(ptMonthName(mk))}</div>
          </div>
          <div>
            <div class="month-eyebrow">Pedidos</div>
            <div class="month-total">${fmtNum(m.total)}</div>
          </div>
        </div>
        ${renderBar(m)}
        ${renderCategoryGrid(m)}
        ${renderSaved(m)}
        ${renderMonthOffenders(m.offenders, 'month')}
        ${renderRastreios(m)}
      </section>`;
  }).join('');

  const consolidatedSection = analysis.sortedMonths.length > 1 ? `
    <section class="month-card consolidated page-break">
      <div class="month-head">
        <div>
          <div class="month-eyebrow">Consolidado</div>
          <div class="month-title">Período completo</div>
        </div>
        <div>
          <div class="month-eyebrow">Pedidos</div>
          <div class="month-total">${fmtNum(analysis.overall.total)}</div>
        </div>
      </div>
      ${renderBar(analysis.overall)}
      ${renderCategoryGrid(analysis.overall)}
      ${renderSaved(analysis.overall, { consolidated: true })}
    </section>` : '';

  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Relatório de Volumetria — ${escapeHTML(filename)}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,300;9..144,400;9..144,500;9..144,600;9..144,700&family=DM+Sans:opsz,wght@9..40,400;9..40,500;9..40,600;9..40,700&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<style>${css}</style>
</head>
<body>
  ${clientName && clientName.trim() ? `
    <div class="client-banner">
      <div class="client-banner-label">Relatório de volumetria preparado para</div>
      <div class="client-banner-name">${escapeHTML(clientName.trim())}</div>
    </div>
  ` : ''}

  <div class="report-header">
    <div class="brand-mark">▦</div>
    <div class="report-header-text">
      <div class="brand-title">Analisador de Volumetria</div>
      <div class="brand-meta">Base OPS · 1ª Tentativa · Transportadoras</div>
    </div>
  </div>

  ${summarySection}

  <h2 class="section-h">${analysis.sortedMonths.length > 1 ? 'Desempenho por mês' : 'Desempenho do período'}</h2>
  <p class="method-note"><strong>Metodologia:</strong> a classificação é aplicada diretamente sobre as datas. Um pedido conta como <strong>no prazo</strong> em qualquer um destes casos: a 1ª tentativa de entrega ocorreu dentro do prazo (mesmo que a entrega final tenha sido posterior); foi entregue até a data prevista; ou ainda está em trânsito e a previsão ainda não venceu. Pedidos entram em <em>fora do prazo</em> ou <em>em atraso</em> quando nenhuma tentativa ocorreu no prazo previsto; encomendas <em>extraviadas</em> ou <em>em devolução</em> contam sempre como problema, independentemente das datas.</p>
  ${monthSections}
  ${consolidatedSection}

  <div class="page-break"></div>
  <h2 class="section-h">Top 3 transportadoras ofensoras</h2>
  ${renderBigOffenders()}

  <footer class="report-footer">
    <div class="footer-partners">
      <div class="footer-partner">
        <img src="${LOGO_MANDAE}" alt="Mandaê" class="footer-logo" />
      </div>
      <div class="footer-partner">
        <img src="${LOGO_ENVIO}" alt="Nuvem Envio" class="footer-logo" />
      </div>
    </div>
    <div class="footer-tagline">Parceiros oficiais</div>
  </footer>

  <script>
    (function() {
      function fire() {
        try { window.focus(); window.print(); } catch(e) {}
      }
      if (document.fonts && document.fonts.ready) {
        Promise.race([
          document.fonts.ready,
          new Promise(function(r) { setTimeout(r, 1500); })
        ]).then(function() { setTimeout(fire, 250); });
      } else {
        setTimeout(fire, 800);
      }
    })();
  </script>
</body>
</html>`;
};

/* ---------- main ---------- */

export default function App() {
  const [data, setData] = useState(null);
  const [error, setError] = useState(null);
  const [filename, setFilename] = useState(null);
  const [clientName, setClientName] = useState('');
  const [loading, setLoading] = useState(false);
  const [dragActive, setDragActive] = useState(false);
  const inputRef = useRef(null);

  const handleFile = useCallback(async (file) => {
    if (!file) return;
    setLoading(true);
    setError(null);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
      const clean = rows.filter((r) => {
        const id = r.codigo_rastreamento;
        if (!id) return false;
        if (String(id).toLowerCase() === 'total') return false;
        if (String(r.status || '').toLowerCase() === 'total') return false;
        if (String(r.transportadora || '').toLowerCase() === 'total') return false;
        return true;
      });
      if (clean.length === 0) throw new Error('Nenhum pedido válido encontrado na planilha.');
      setData(clean);
      setFilename(file.name);
      // Se a planilha inteira for de um único cliente, pré-preenche o nome no relatório.
      // O usuário ainda pode editar/limpar o campo manualmente.
      const uniqueClients = [
        ...new Set(clean.map((r) => String(r.nome_comercial ?? '').trim()).filter(Boolean)),
      ];
      if (uniqueClients.length === 1) setClientName(uniqueClients[0]);
    } catch (e) {
      setError('Não foi possível ler o arquivo: ' + e.message);
      setData(null);
    } finally {
      setLoading(false);
    }
  }, []);

  const onDrop = useCallback(
    (e) => {
      e.preventDefault();
      setDragActive(false);
      const f = e.dataTransfer.files?.[0];
      if (f) handleFile(f);
    },
    [handleFile]
  );

  const reset = () => {
    setData(null);
    setError(null);
    setFilename(null);
    setClientName('');
    if (inputRef.current) inputRef.current.value = '';
  };

  const analysis = useMemo(() => {
    if (!data) return null;

    const enriched = data.map((r) => {
      const dataEnvio = parseDateBR(r.data_envio ?? r.data_hora_envio);
      const dataPrim = parseDateBR(r.data_primeira_tentativa_entrega);
      const dataPrev = parseDateBR(r.data_previsao_entrega_cliente);
      const dataEnt = parseDateBR(r.data_entrega);
      const status = String(r.status || '').trim();
      // "Salvos pela 1ª tentativa": entregue + 1ª tentativa <= prazo + entrega > prazo
      const savedByFirstAttempt =
        status === 'Entregue' &&
        dataPrim != null && dataPrev != null && dataEnt != null &&
        dataPrim.getTime() <= dataPrev.getTime() &&
        dataEnt.getTime() > dataPrev.getTime();
      return {
        ...r,
        _category: classify(r),
        _monthKey: monthKeyOf(dataEnvio),
        _carrier: String(r.transportadora || 'Não informada').trim() || 'Não informada',
        _savedByFirstAttempt: savedByFirstAttempt,
      };
    });

    const blankCats = () => Object.fromEntries(CATEGORIES.map((c) => [c.key, 0]));

    const months = {};
    const overall = {
      total: enriched.length,
      categories: blankCats(),
      byCarrier: {},
      savedRows: [],
    };

    enriched.forEach((r) => {
      // overall
      overall.categories[r._category]++;
      if (r._savedByFirstAttempt) overall.savedRows.push(r);
      if (!overall.byCarrier[r._carrier]) {
        overall.byCarrier[r._carrier] = { total: 0, categories: blankCats() };
      }
      overall.byCarrier[r._carrier].total++;
      overall.byCarrier[r._carrier].categories[r._category]++;

      // monthly
      if (!r._monthKey) return;
      if (!months[r._monthKey]) {
        months[r._monthKey] = {
          total: 0,
          categories: blankCats(),
          byCarrier: {},
          rows: Object.fromEntries(PROBLEM_ROW_KEYS.map((k) => [k, []])),
          savedRows: [],
        };
      }
      months[r._monthKey].total++;
      months[r._monthKey].categories[r._category]++;
      if (PROBLEM_ROW_KEYS.includes(r._category)) {
        months[r._monthKey].rows[r._category].push(r);
      }
      if (r._savedByFirstAttempt) months[r._monthKey].savedRows.push(r);
      if (!months[r._monthKey].byCarrier[r._carrier]) {
        months[r._monthKey].byCarrier[r._carrier] = { total: 0, categories: blankCats() };
      }
      months[r._monthKey].byCarrier[r._carrier].total++;
      months[r._monthKey].byCarrier[r._carrier].categories[r._category]++;
    });

    const buildOffenders = (totalUniverse, byCarrier) =>
      Object.entries(byCarrier)
        .map(([name, info]) => {
          const problems = CATEGORIES
            .filter((c) => c.isProblem)
            .reduce((s, c) => s + info.categories[c.key], 0);
          return {
            name,
            ...info,
            problems,
            problemRate: info.total > 0 ? problems / info.total : 0,
            shareOfTotal: totalUniverse > 0 ? problems / totalUniverse : 0,
            volumeShare: totalUniverse > 0 ? info.total / totalUniverse : 0,
          };
        })
        .filter((o) => o.problems > 0)
        .sort((a, b) => b.problems - a.problems)
        .slice(0, 3);

    const offenders = buildOffenders(overall.total, overall.byCarrier);

    Object.keys(months).forEach((mk) => {
      months[mk].offenders = buildOffenders(months[mk].total, months[mk].byCarrier);
    });

    const sortedMonths = Object.keys(months).sort();
    return { overall, months, sortedMonths, offenders };
  }, [data]);

  const exportPDF = useCallback(() => {
    if (!analysis) return;
    const html = buildReportHTML(analysis, filename || 'Base_OPS.xlsx', formatTodayBR(), clientName);
    const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const stamp = formatTodayBR().replace(/\//g, '-');
    const safeClient = (clientName || '').trim().replace(/[^\w\u00C0-\u017F\s-]/g, '').slice(0, 40);
    const fname = safeClient
      ? `Relatorio_Volumetria_${safeClient}_${stamp}.html`
      : `Relatorio_Volumetria_${stamp}.html`;

    let opened = null;
    try {
      opened = window.open(url, '_blank', 'noopener,noreferrer');
    } catch (e) {
      opened = null;
    }

    if (opened) {
      setTimeout(() => URL.revokeObjectURL(url), 60000);
      return;
    }

    // Fallback: trigger download (popup blocked or window.open unavailable in sandbox)
    const a = document.createElement('a');
    a.href = url;
    a.download = fname;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }, [analysis, filename, clientName]);

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,300;9..144,400;9..144,500;9..144,600;9..144,700&family=DM+Sans:opsz,wght@9..40,400;9..40,500;9..40,600;9..40,700&family=JetBrains+Mono:wght@400;500;600&display=swap');
        .ff-display { font-family: 'Fraunces', Georgia, serif; font-optical-sizing: auto; }
        .ff-body { font-family: 'DM Sans', system-ui, sans-serif; }
        .ff-mono { font-family: 'JetBrains Mono', ui-monospace, monospace; }
        .grain {
          background-image:
            radial-gradient(circle at 20% 10%, rgba(180,83,9,0.04), transparent 40%),
            radial-gradient(circle at 80% 90%, rgba(28,28,46,0.04), transparent 40%);
        }
        @keyframes fadeUp { from { opacity:0; transform: translateY(8px); } to { opacity:1; transform:none; } }
        .anim-fade { animation: fadeUp .4s ease-out both; }

        /* ---------- Print stylesheet ---------- */
        @media print {
          @page { size: A4; margin: 14mm 12mm; }
          * {
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
            color-adjust: exact !important;
          }
          html, body { background: #ffffff !important; }
          .grain { background-image: none !important; }
          /* Force-show every collapsed rastreio list and undo height clipping */
          .rastreio-list { display: block !important; max-height: none !important; overflow: visible !important; }
          /* Page-break controls */
          .avoid-break { break-inside: avoid; page-break-inside: avoid; }
          .month-card { break-inside: auto; page-break-inside: auto; }
          .month-card + .month-card { break-before: page; page-break-before: always; }
          /* Avoid orphan headers */
          h2, .month-header { break-after: avoid; page-break-after: avoid; }
          /* Animations off in print */
          .anim-fade { animation: none !important; }
        }
      `}</style>

      <div className="ff-body min-h-screen bg-[#faf7f2] text-[#1c1c2e] grain">
        {/* Header */}
        <header className="border-b border-[#1c1c2e]/10">
          <div className="max-w-6xl mx-auto px-6 py-6 flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 bg-[#1c1c2e] rounded-sm flex items-center justify-center">
                <Package className="w-5 h-5 text-[#faf7f2]" strokeWidth={1.5} />
              </div>
              <div>
                <div className="ff-display text-xl leading-none tracking-tight">Analisador de Volumetria</div>
                <div className="text-[11px] uppercase tracking-[0.2em] text-[#1c1c2e]/50 mt-1">Base OPS · 1ª Tentativa · Transportadoras</div>
              </div>
            </div>
            {data && (
              <div className="flex items-center gap-2 print:hidden">
                <button
                  onClick={exportPDF}
                  className="flex items-center gap-2 text-sm px-4 py-2 bg-[#1c1c2e] text-[#faf7f2] hover:bg-[#b45309] transition rounded-sm"
                >
                  <Printer className="w-4 h-4" />
                  Exportar PDF
                </button>
                <button
                  onClick={reset}
                  className="flex items-center gap-2 text-sm px-4 py-2 border border-[#1c1c2e]/20 hover:border-[#1c1c2e]/60 hover:bg-[#1c1c2e] hover:text-[#faf7f2] transition rounded-sm"
                >
                  <RotateCcw className="w-4 h-4" />
                  Nova análise
                </button>
              </div>
            )}
          </div>
        </header>

        {/* Tableau source link bar (sempre visível na tela, escondido no PDF) */}
        <div className="border-b border-[#1c1c2e]/10 bg-[#1c1c2e]/[0.02] print:hidden">
          <div className="max-w-6xl mx-auto px-6 py-2.5 flex items-center gap-2 text-xs flex-wrap">
            <span className="text-[#1c1c2e]/55 uppercase tracking-wider text-[10px]">Fonte da planilha:</span>
            <a
              href={TABLEAU_URL}
              target="_blank"
              rel="noopener noreferrer"
              className="ff-mono text-[#b45309] hover:underline break-all"
            >
              {TABLEAU_URL}
            </a>
          </div>
        </div>

        <main className="max-w-6xl mx-auto px-6 py-10">
          {/* Empty state */}
          {!data && (
            <section className="anim-fade">
              <div className="text-center mb-10">
                <div className="text-[11px] uppercase tracking-[0.3em] text-[#b45309] mb-4">Etapa 01</div>
                <h1 className="ff-display text-5xl md:text-6xl font-light leading-[1.05] tracking-tight max-w-3xl mx-auto">
                  Envie a planilha de <em className="text-[#b45309] not-italic font-normal">Base OPS</em> e receba o relatório de desempenho
                </h1>
                <p className="mt-6 text-[#1c1c2e]/70 max-w-xl mx-auto">
                  A planilha precisa seguir o modelo padrão (com colunas <span className="ff-mono text-xs bg-[#1c1c2e]/5 px-1.5 py-0.5 rounded">codigo_rastreamento</span>, <span className="ff-mono text-xs bg-[#1c1c2e]/5 px-1.5 py-0.5 rounded">status</span>, <span className="ff-mono text-xs bg-[#1c1c2e]/5 px-1.5 py-0.5 rounded">transportadora</span>, <span className="ff-mono text-xs bg-[#1c1c2e]/5 px-1.5 py-0.5 rounded">data_hora_envio</span>, <span className="ff-mono text-xs bg-[#1c1c2e]/5 px-1.5 py-0.5 rounded">data_primeira_tentativa_entrega</span>, <span className="ff-mono text-xs bg-[#1c1c2e]/5 px-1.5 py-0.5 rounded">data_previsao_entrega_cliente</span> e <span className="ff-mono text-xs bg-[#1c1c2e]/5 px-1.5 py-0.5 rounded">data_entrega</span>).
                </p>
                <p className="mt-4 text-[13px] text-[#1c1c2e]/55 max-w-xl mx-auto leading-relaxed">
                  <strong className="text-[#1c1c2e]/75">Critério de prazo:</strong> pedidos cuja 1ª tentativa de entrega ocorreu dentro do prazo são considerados <em>no prazo</em>, mesmo que a entrega final tenha sido posterior. Pedidos só entram em <em>fora do prazo</em> ou <em>em atraso</em> se nenhuma tentativa ocorreu dentro do prazo previsto.
                </p>
              </div>

              <div
                onDragOver={(e) => { e.preventDefault(); setDragActive(true); }}
                onDragLeave={() => setDragActive(false)}
                onDrop={onDrop}
                onClick={() => inputRef.current?.click()}
                className={`cursor-pointer border-2 border-dashed transition-all rounded-sm py-16 px-8 text-center ${
                  dragActive
                    ? 'border-[#b45309] bg-[#b45309]/5'
                    : 'border-[#1c1c2e]/25 hover:border-[#1c1c2e]/50 bg-white/40'
                }`}
              >
                <input
                  ref={inputRef}
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  className="hidden"
                  onChange={(e) => handleFile(e.target.files?.[0])}
                />
                <div className="flex flex-col items-center gap-4">
                  <div className="w-14 h-14 rounded-full bg-[#1c1c2e] flex items-center justify-center">
                    <Upload className="w-6 h-6 text-[#faf7f2]" strokeWidth={1.5} />
                  </div>
                  <div>
                    <div className="ff-display text-2xl">Arraste sua planilha aqui</div>
                    <div className="text-sm text-[#1c1c2e]/60 mt-1">ou clique para selecionar · .xlsx, .xls ou .csv</div>
                  </div>
                </div>
              </div>

              {loading && (
                <div className="mt-6 text-center text-sm text-[#1c1c2e]/60">Processando arquivo…</div>
              )}
              {error && (
                <div className="mt-6 border border-red-200 bg-red-50 text-red-800 px-4 py-3 rounded-sm text-sm">
                  {error}
                </div>
              )}
            </section>
          )}

          {/* Results */}
          {analysis && (
            <div className="space-y-12 anim-fade">
              {/* Summary banner */}
              <section>
                <div className="flex items-baseline gap-3 mb-4 flex-wrap">
                  <FileSpreadsheet className="w-4 h-4 text-[#1c1c2e]/50" />
                  <span className="ff-mono text-xs text-[#1c1c2e]/60">{filename}</span>
                  <span className="text-[#1c1c2e]/30">·</span>
                  <span className="text-xs text-[#1c1c2e]/60 flex items-center gap-1">
                    <Calendar className="w-3 h-3" /> Análise em {formatTodayBR()}
                  </span>
                </div>

                {/* Client name input */}
                <div className="mb-6 print:hidden">
                  <label htmlFor="client-name" className="block text-[11px] uppercase tracking-[0.2em] text-[#1c1c2e]/50 mb-1.5">
                    Cliente
                  </label>
                  <input
                    id="client-name"
                    type="text"
                    value={clientName}
                    onChange={(e) => setClientName(e.target.value)}
                    placeholder="Digite o nome do cliente para incluir no relatório"
                    className="w-full max-w-2xl px-4 py-2.5 text-base bg-white border border-[#1c1c2e]/20 rounded-sm focus:outline-none focus:border-[#b45309] focus:ring-1 focus:ring-[#b45309] transition placeholder:text-[#1c1c2e]/35"
                    autoComplete="off"
                  />
                  <p className="text-[11px] text-[#1c1c2e]/50 mt-1.5">
                    Aparece no cabeçalho do relatório quando você exportar para PDF.
                  </p>
                </div>

                <div className="grid md:grid-cols-3 gap-6 border-t border-[#1c1c2e]/15 pt-6">
                  <div>
                    <div className="text-[11px] uppercase tracking-[0.2em] text-[#1c1c2e]/50">Total de encomendas</div>
                    <div className="ff-display text-6xl font-light tracking-tight mt-1">{fmtNum(analysis.overall.total)}</div>
                  </div>
                  <div>
                    <div className="text-[11px] uppercase tracking-[0.2em] text-[#1c1c2e]/50">Período coberto</div>
                    <div className="ff-display text-2xl mt-2 leading-snug">
                      {analysis.sortedMonths.length === 0
                        ? '—'
                        : analysis.sortedMonths.length === 1
                        ? ptMonthName(analysis.sortedMonths[0])
                        : `${ptMonthName(analysis.sortedMonths[0])} → ${ptMonthName(analysis.sortedMonths[analysis.sortedMonths.length - 1])}`}
                    </div>
                    <div className="text-xs text-[#1c1c2e]/60 mt-1">{analysis.sortedMonths.length} {analysis.sortedMonths.length === 1 ? 'mês' : 'meses'}</div>
                  </div>
                  <div>
                    <div className="text-[11px] uppercase tracking-[0.2em] text-[#1c1c2e]/50">Transportadoras</div>
                    <div className="ff-display text-2xl mt-2">{Object.keys(analysis.overall.byCarrier).length}</div>
                    <div className="text-xs text-[#1c1c2e]/60 mt-1">parceiros operando no período</div>
                  </div>
                </div>
              </section>

              {/* Monthly breakdown */}
              <section>
                <div className="flex items-baseline justify-between mb-3 border-b border-[#1c1c2e]/15 pb-3">
                  <h2 className="ff-display text-3xl">
                    {analysis.sortedMonths.length > 1 ? 'Desempenho por mês' : 'Desempenho do período'}
                  </h2>
                  <span className="text-[11px] uppercase tracking-[0.2em] text-[#1c1c2e]/50">Etapa 02</span>
                </div>
                <p className="text-xs text-[#1c1c2e]/60 mb-6 max-w-3xl leading-relaxed">
                  <strong className="text-[#1c1c2e]/80">Metodologia:</strong> a classificação é aplicada diretamente sobre as datas. Um pedido conta como <strong>no prazo</strong> em qualquer um destes casos: a 1ª tentativa de entrega ocorreu dentro do prazo (mesmo que a entrega final tenha sido posterior); foi entregue até a data prevista; ou ainda está em trânsito e a previsão ainda não venceu. Pedidos entram em <em>fora do prazo</em> ou <em>em atraso</em> quando nenhuma tentativa ocorreu no prazo previsto; encomendas <em>extraviadas</em> ou <em>em devolução</em> contam sempre como problema, independentemente das datas.
                </p>

                <div className="space-y-8">
                  {analysis.sortedMonths.map((mk) => (
                    <MonthCard
                      key={mk}
                      monthLabel={ptMonthName(mk)}
                      data={analysis.months[mk]}
                    />
                  ))}
                  {analysis.sortedMonths.length > 1 && (
                    <MonthCard
                      monthLabel="Consolidado do período"
                      data={analysis.overall}
                      consolidated
                    />
                  )}
                </div>
              </section>

              {/* Top offenders */}
              <section>
                <div className="flex items-baseline justify-between mb-6 border-b border-[#1c1c2e]/15 pb-3">
                  <h2 className="ff-display text-3xl flex items-center gap-3">
                    <TrendingDown className="w-7 h-7 text-[#b45309]" strokeWidth={1.5} />
                    Top 3 transportadoras ofensoras
                  </h2>
                  <span className="text-[11px] uppercase tracking-[0.2em] text-[#1c1c2e]/50">Etapa 03</span>
                </div>

                {analysis.offenders.length === 0 ? (
                  <div className="border border-emerald-200 bg-emerald-50 text-emerald-900 px-5 py-6 rounded-sm">
                    Nenhuma transportadora apresentou ocorrências no período. 🎉
                  </div>
                ) : (
                  <div className="grid md:grid-cols-3 gap-5">
                    {analysis.offenders.map((o, idx) => (
                      <OffenderCard
                        key={o.name}
                        rank={idx + 1}
                        offender={o}
                        totalOrders={analysis.overall.total}
                      />
                    ))}
                  </div>
                )}

                <p className="text-xs text-[#1c1c2e]/60 mt-5 leading-relaxed max-w-3xl">
                  <strong>Critério:</strong> ranking pelo número absoluto de ocorrências (entregas fora do prazo, em atraso, extravios e devoluções). A porcentagem em relação ao total considera o universo de {fmtNum(analysis.overall.total)} pedidos do período.
                </p>
              </section>
            </div>
          )}
        </main>

        <footer className="max-w-6xl mx-auto px-6 py-8 mt-10 border-t border-[#1c1c2e]/10 print:hidden">
          <div className="flex flex-col items-center gap-2 mb-5">
            <div className="flex items-center justify-center gap-4">
              <div className="w-16 h-16 rounded-md overflow-hidden border border-[#1c1c2e]/12 bg-white flex items-center justify-center">
                <img src={LOGO_MANDAE} alt="Mandaê" className="w-full h-full object-cover" />
              </div>
              <div className="w-16 h-16 rounded-md overflow-hidden border border-[#1c1c2e]/12 bg-white flex items-center justify-center">
                <img src={LOGO_ENVIO} alt="Nuvem Envio" className="w-full h-full object-cover" />
              </div>
            </div>
            <div className="text-[10px] uppercase tracking-[0.3em] text-[#1c1c2e]/50">
              Parceiros oficiais
            </div>
          </div>
          <div className="text-xs text-[#1c1c2e]/50 flex items-center justify-between">
            <span>Análise gerada localmente · seus dados não saem do navegador</span>
            <span className="ff-mono">v3.0 · OPS</span>
          </div>
        </footer>
      </div>
    </>
  );
}

/* ---------- Month card ---------- */

function MonthCard({ monthLabel, data, consolidated = false }) {
  const total = data.total;
  const [openKey, setOpenKey] = useState({});
  const [copiedKey, setCopiedKey] = useState(null);

  const toggle = (k) => setOpenKey((p) => ({ ...p, [k]: !p[k] }));

  const copy = (k, rows) => {
    const codes = rows.map((r) => r.codigo_rastreamento || '').filter(Boolean).join('\n');
    if (!codes) return;
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(codes).then(() => {
        setCopiedKey(k);
        setTimeout(() => setCopiedKey((cur) => (cur === k ? null : cur)), 1600);
      });
    }
  };

  return (
    <div className={`month-card border ${consolidated ? 'border-[#1c1c2e]/40 bg-[#1c1c2e] text-[#faf7f2] print:bg-white print:text-[#1c1c2e] print:border-[#1c1c2e]/40' : 'border-[#1c1c2e]/15 bg-white/60'} rounded-sm overflow-hidden`}>
      <div className="month-header avoid-break px-6 py-5 flex items-baseline justify-between border-b border-current/10">
        <div>
          <div className={`text-[11px] uppercase tracking-[0.25em] ${consolidated ? 'text-[#faf7f2]/60 print:text-[#1c1c2e]/50' : 'text-[#1c1c2e]/50'}`}>
            {consolidated ? 'Consolidado' : 'Mês'}
          </div>
          <div className="ff-display text-3xl mt-1">{monthLabel}</div>
        </div>
        <div className="text-right">
          <div className={`text-[11px] uppercase tracking-[0.25em] ${consolidated ? 'text-[#faf7f2]/60 print:text-[#1c1c2e]/50' : 'text-[#1c1c2e]/50'}`}>
            Pedidos
          </div>
          <div className="ff-display text-4xl font-light mt-1">{fmtNum(total)}</div>
        </div>
      </div>

      {/* Stacked bar */}
      <div className="px-6 pt-5">
        <div className="h-2 w-full flex rounded-full overflow-hidden bg-current/10">
          {CATEGORIES.map((c) => {
            const v = data.categories[c.key];
            const pct = total > 0 ? (v / total) * 100 : 0;
            if (pct === 0) return null;
            return (
              <div
                key={c.key}
                style={{ width: `${pct}%`, backgroundColor: COLOR[c.color].hex }}
                title={`${c.label}: ${v} (${fmtPct(v / total)})`}
              />
            );
          })}
        </div>
      </div>

      {/* Categories grid */}
      <div className="px-6 pb-6 pt-5 grid grid-cols-1 sm:grid-cols-3 gap-3">
        {CATEGORIES.map((c) => {
          const v = data.categories[c.key];
          const pct = total > 0 ? v / total : 0;
          const Icon = c.icon;
          const cl = COLOR[c.color];
          const showSavedNote = c.key === 'no_prazo' && data.savedRows && data.savedRows.length > 0;
          return (
            <div
              key={c.key}
              className={`avoid-break ${consolidated ? 'bg-white/5 border-white/10 print:bg-white print:border-[#1c1c2e]/15' : `${cl.soft} ${cl.border}`} border rounded-sm px-4 py-3 flex items-center gap-3`}
            >
              <div className={`w-9 h-9 rounded-full flex items-center justify-center ${cl.solid} text-white shrink-0`}>
                <Icon className="w-4 h-4" strokeWidth={2} />
              </div>
              <div className="flex-1 min-w-0">
                <div className={`text-[11px] uppercase tracking-wider ${consolidated ? 'text-[#faf7f2]/60 print:text-[#1c1c2e]/60' : 'text-[#1c1c2e]/60'}`}>
                  {c.short}
                </div>
                <div className="flex items-baseline gap-2 mt-0.5">
                  <span className="ff-mono text-lg font-medium">{fmtNum(v)}</span>
                  <span className={`ff-mono text-xs ${consolidated ? 'text-[#faf7f2]/70 print:text-[#1c1c2e]/60' : 'text-[#1c1c2e]/60'}`}>
                    {fmtPct(pct)}
                  </span>
                </div>
                {showSavedNote && (
                  <div className={`text-[10px] mt-0.5 leading-tight ${consolidated ? 'text-[#faf7f2]/55 print:text-[#1c1c2e]/55' : 'text-[#1c1c2e]/55'}`}>
                    inclui {fmtNum(data.savedRows.length)} com 1ª tent. no prazo
                  </div>
                )}
              </div>
            </div>
          );
        })}
      </div>

      {/* Saved by first attempt */}
      {data.savedRows && data.savedRows.length > 0 && (
        <div className={`border-t ${consolidated ? 'border-white/15' : 'border-[#1c1c2e]/10'} px-6 pt-5 pb-6 ${consolidated ? '' : 'bg-emerald-50/40'}`}>
          <div className="flex items-center gap-2 mb-2">
            <Sparkles className={`w-4 h-4 ${consolidated ? 'text-emerald-300 print:text-emerald-700' : 'text-emerald-700'}`} strokeWidth={2} />
            <div className={`text-[11px] uppercase tracking-[0.25em] ${consolidated ? 'text-[#faf7f2]/60 print:text-[#1c1c2e]/60' : 'text-[#1c1c2e]/60'}`}>
              Entregas tardias com 1ª tentativa no prazo
            </div>
          </div>

          <div className="flex items-baseline gap-3 mb-3">
            <span className={`ff-display text-4xl font-light leading-none ${consolidated ? '' : 'text-emerald-700'}`}>
              {fmtNum(data.savedRows.length)}
            </span>
            <span className={`text-xs ${consolidated ? 'text-[#faf7f2]/65 print:text-[#1c1c2e]/60' : 'text-[#1c1c2e]/60'} leading-snug max-w-md`}>
              pedido{data.savedRows.length > 1 ? 's' : ''} cuja 1ª tentativa de entrega ocorreu dentro do prazo, mas a entrega final foi posterior à previsão. Sob a regra de OTD ajustado, contam como <strong>no prazo</strong>.
            </span>
          </div>

          <SavedRowsList
            rows={data.savedRows}
            instanceKey={`saved:${consolidated ? 'consolidated' : data.total}`}
            openKey={openKey}
            toggle={toggle}
            copiedKey={copiedKey}
            copy={copy}
          />

          <p className={`text-[10.5px] mt-3 leading-relaxed ${consolidated ? 'text-[#faf7f2]/55 print:text-[#1c1c2e]/55' : 'text-[#1c1c2e]/55'}`}>
            <strong>Sobre as causas:</strong> a Base OPS não traz coluna de descrição de evento. As datas mostradas (1ª tentativa, previsão, entrega) e o intervalo entre 1ª tentativa e entrega final são as melhores pistas disponíveis: gap pequeno (1–2 dias) sugere reentrega rápida; gap maior indica investigação cruzada com a base de eventos.
          </p>
        </div>
      )}

      {/* Monthly offenders */}
      {!consolidated && data.offenders && data.offenders.length > 0 && (
        <div className="border-t border-[#1c1c2e]/10 px-6 pt-5 pb-6 bg-[#faf7f2]/40">
          <div className="flex items-center gap-2 mb-3">
            <TrendingDown className="w-4 h-4 text-[#b45309]" strokeWidth={2} />
            <div className="text-[11px] uppercase tracking-[0.25em] text-[#1c1c2e]/60">
              Top {data.offenders.length} ofensoras do mês
            </div>
          </div>
          <div className="grid md:grid-cols-3 gap-2">
            {data.offenders.map((o, idx) => {
              const issues = CATEGORIES
                .filter((c) => c.isProblem)
                .map((c) => ({ ...c, count: o.categories[c.key] }))
                .filter((c) => c.count > 0)
                .sort((a, b) => b.count - a.count);
              const headline = issues[0];
              return (
                <div key={o.name} className="border border-[#1c1c2e]/15 bg-white rounded-sm px-3 py-3">
                  <div className="flex items-baseline gap-2 mb-1.5">
                    <span className="ff-display text-2xl font-light text-[#b45309] leading-none">
                      {String(idx + 1).padStart(2, '0')}
                    </span>
                    <span className="ff-display text-base leading-tight truncate flex-1" title={o.name}>
                      {o.name}
                    </span>
                  </div>
                  <div className="text-xs text-[#1c1c2e]/70 leading-snug">
                    <span className="ff-mono font-medium text-[#1c1c2e]">{fmtNum(o.problems)}</span> ocorrência{o.problems > 1 ? 's' : ''}
                    {' · '}
                    <span className="ff-mono">{fmtPct(o.problemRate)}</span> dos próprios pedidos
                  </div>
                  <div className="text-xs text-[#1c1c2e]/70 mt-0.5">
                    <span className="ff-mono">{fmtPct(o.shareOfTotal)}</span> do mês · vol. {fmtNum(o.total)}
                  </div>
                  {headline && (
                    <div className="mt-2 pt-2 border-t border-[#1c1c2e]/10 flex items-center gap-1.5">
                      <span
                        className="w-1.5 h-1.5 rounded-full shrink-0"
                        style={{ backgroundColor: COLOR[headline.color].hex }}
                      />
                      <span className="text-[11px] text-[#1c1c2e]/70 truncate">
                        Principal: {headline.label.toLowerCase()} ({fmtNum(headline.count)})
                      </span>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Monthly tracking lists */}
      {!consolidated && data.rows && PROBLEM_ROW_KEYS.some((k) => data.rows[k].length > 0) && (
        <div className="border-t border-[#1c1c2e]/10 px-6 pt-5 pb-6">
          <div className="flex items-center gap-2 mb-3">
            <Package className="w-4 h-4 text-[#1c1c2e]/50" strokeWidth={2} />
            <div className="text-[11px] uppercase tracking-[0.25em] text-[#1c1c2e]/60">
              Rastreios com ocorrências
            </div>
          </div>
          <div className="space-y-2">
            {PROBLEM_ROW_KEYS.map((key) => {
              const cat = CATEGORIES.find((c) => c.key === key);
              const rows = data.rows[key] || [];
              if (rows.length === 0) return null;
              const isOpen = !!openKey[key];
              const cl = COLOR[cat.color];
              return (
                <div key={key} className="avoid-break border border-[#1c1c2e]/15 bg-white rounded-sm overflow-hidden">
                  <div className="px-4 py-3 flex items-center justify-between gap-3">
                    <button
                      onClick={() => toggle(key)}
                      className="flex items-center gap-3 flex-1 text-left min-w-0"
                    >
                      <span
                        className="w-2.5 h-2.5 rounded-full shrink-0"
                        style={{ backgroundColor: cl.hex }}
                      />
                      <span className="text-sm font-medium truncate">{cat.label}</span>
                      <span className="ff-mono text-xs text-[#1c1c2e]/60 shrink-0">
                        {fmtNum(rows.length)} rastreio{rows.length > 1 ? 's' : ''}
                      </span>
                      <ChevronDown
                        className={`w-4 h-4 text-[#1c1c2e]/50 transition-transform shrink-0 print:hidden ${
                          isOpen ? 'rotate-180' : ''
                        }`}
                        strokeWidth={2}
                      />
                    </button>
                    <button
                      onClick={() => copy(key, rows)}
                      title="Copiar rastreios"
                      className={`flex items-center gap-1 text-xs px-2 py-1 rounded transition shrink-0 print:hidden ${
                        copiedKey === key
                          ? 'bg-emerald-50 text-emerald-700'
                          : 'hover:bg-[#1c1c2e]/5 text-[#1c1c2e]/70'
                      }`}
                    >
                      {copiedKey === key ? (
                        <>
                          <Check className="w-3.5 h-3.5" strokeWidth={2.5} />
                          <span>Copiado</span>
                        </>
                      ) : (
                        <>
                          <Copy className="w-3.5 h-3.5" strokeWidth={2} />
                          <span className="hidden sm:inline">Copiar</span>
                        </>
                      )}
                    </button>
                  </div>
                  <div
                    className={`rastreio-list border-t border-[#1c1c2e]/10 max-h-80 overflow-y-auto ${
                      isOpen ? 'block' : 'hidden'
                    }`}
                  >
                    {rows.map((r, i) => (
                      <TrackingRow key={i} row={r} category={key} accent={cl.hex} />
                    ))}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------- Tracking row ---------- */

function TrackingRow({ row, category, accent }) {
  const code = row.codigo_rastreamento || '—';
  const carrier = row._carrier || row.transportadora || '—';
  const dataPrev = parseDateBR(row.data_previsao_entrega_cliente);
  const dataEnt = parseDateBR(row.data_entrega);
  const dataEnvio = parseDateBR(row.data_hora_envio);
  const dataPrim = parseDateBR(row.data_primeira_tentativa_entrega);

  let detail = null;
  let badge = null;
  let badgeClass = 'bg-[#1c1c2e]/5 text-[#1c1c2e]/80';

  if (category === 'fora_prazo') {
    const days = daysBetween(dataEnt, dataPrev);
    detail = (
      <>
        prev. <span className="ff-mono">{fmtShortDate(dataPrev)}</span>
        {' → entrega '}
        <span className="ff-mono">{fmtShortDate(dataEnt)}</span>
        {dataPrim && (
          <>
            {' · 1ª tent. '}
            <span className="ff-mono">{fmtShortDate(dataPrim)}</span>
          </>
        )}
      </>
    );
    if (days != null) {
      badge = `+${days}d`;
      badgeClass = 'bg-amber-50 text-amber-800 border border-amber-200';
    }
  } else if (category === 'em_atraso') {
    const days = daysBetween(today, dataPrev);
    detail = (
      <>
        envio <span className="ff-mono">{fmtShortDate(dataEnvio)}</span>
        {' · prev. '}
        <span className="ff-mono">{fmtShortDate(dataPrev)}</span>
        {dataPrim && (
          <>
            {' · 1ª tent. '}
            <span className="ff-mono">{fmtShortDate(dataPrim)}</span>
          </>
        )}
      </>
    );
    if (days != null) {
      badge = `+${days}d`;
      badgeClass = 'bg-red-50 text-red-800 border border-red-200';
    }
  }

  return (
    <div className="px-4 py-2.5 border-b border-[#1c1c2e]/5 last:border-b-0 flex flex-wrap items-center gap-x-3 gap-y-1 hover:bg-[#faf7f2]/50 transition">
      <span
        className="w-1 h-4 rounded-full shrink-0"
        style={{ backgroundColor: accent }}
      />
      <span className="ff-mono text-sm font-medium text-[#1c1c2e]">{code}</span>
      <span className="text-xs text-[#1c1c2e]/60 truncate">{carrier}</span>
      <span className="text-xs text-[#1c1c2e]/60 flex-1 min-w-0 truncate">{detail}</span>
      {badge && (
        <span className={`ff-mono text-[11px] px-2 py-0.5 rounded ${badgeClass} shrink-0`}>
          {badge}
        </span>
      )}
    </div>
  );
}

/* ---------- Saved rows list (entregas tardias com 1ª tentativa no prazo) ---------- */

function SavedRowsList({ rows, instanceKey, openKey, toggle, copiedKey, copy }) {
  const isOpen = !!openKey[instanceKey];
  return (
    <div className="avoid-break border border-emerald-200 bg-white rounded-sm overflow-hidden">
      <div className="px-4 py-3 flex items-center justify-between gap-3">
        <button
          onClick={() => toggle(instanceKey)}
          className="flex items-center gap-3 flex-1 text-left min-w-0"
        >
          <span className="w-2.5 h-2.5 rounded-full shrink-0 bg-emerald-600" />
          <span className="text-sm font-medium truncate">
            {rows.length} rastreio{rows.length > 1 ? 's' : ''} com 1ª tentativa no prazo
          </span>
          <ChevronDown
            className={`w-4 h-4 text-[#1c1c2e]/50 transition-transform shrink-0 print:hidden ${
              isOpen ? 'rotate-180' : ''
            }`}
            strokeWidth={2}
          />
        </button>
        <button
          onClick={() => copy(instanceKey, rows)}
          title="Copiar rastreios"
          className={`flex items-center gap-1 text-xs px-2 py-1 rounded transition shrink-0 print:hidden ${
            copiedKey === instanceKey
              ? 'bg-emerald-50 text-emerald-700'
              : 'hover:bg-[#1c1c2e]/5 text-[#1c1c2e]/70'
          }`}
        >
          {copiedKey === instanceKey ? (
            <>
              <Check className="w-3.5 h-3.5" strokeWidth={2.5} />
              <span>Copiado</span>
            </>
          ) : (
            <>
              <Copy className="w-3.5 h-3.5" strokeWidth={2} />
              <span className="hidden sm:inline">Copiar</span>
            </>
          )}
        </button>
      </div>
      <div
        className={`rastreio-list border-t border-emerald-100 max-h-80 overflow-y-auto ${
          isOpen ? 'block' : 'hidden'
        }`}
      >
        {rows.map((r, i) => (
          <SavedRow key={i} row={r} />
        ))}
      </div>
    </div>
  );
}

function SavedRow({ row }) {
  const code = row.codigo_rastreamento || '—';
  const carrier = row._carrier || row.transportadora || '—';
  const dataPrim = parseDateBR(row.data_primeira_tentativa_entrega);
  const dataPrev = parseDateBR(row.data_previsao_entrega_cliente);
  const dataEnt = parseDateBR(row.data_entrega);
  const gapDays = daysBetween(dataEnt, dataPrim);
  const lateDays = daysBetween(dataEnt, dataPrev);

  return (
    <div className="px-4 py-2.5 border-b border-emerald-50 last:border-b-0 flex flex-wrap items-center gap-x-3 gap-y-1 hover:bg-emerald-50/40 transition">
      <span className="w-1 h-4 rounded-full shrink-0 bg-emerald-600" />
      <span className="ff-mono text-sm font-medium text-[#1c1c2e]">{code}</span>
      <span className="text-xs text-[#1c1c2e]/60 truncate">{carrier}</span>
      <span className="text-xs text-[#1c1c2e]/60 flex-1 min-w-0 truncate">
        1ª tent. <span className="ff-mono">{fmtShortDate(dataPrim)}</span>
        {' · prev. '}
        <span className="ff-mono">{fmtShortDate(dataPrev)}</span>
        {' · entrega '}
        <span className="ff-mono">{fmtShortDate(dataEnt)}</span>
      </span>
      {gapDays != null && (
        <span
          className="ff-mono text-[11px] px-2 py-0.5 rounded bg-emerald-50 text-emerald-800 border border-emerald-200 shrink-0"
          title="Intervalo entre 1ª tentativa e entrega final"
        >
          gap {gapDays}d
        </span>
      )}
      {lateDays != null && lateDays > 0 && (
        <span
          className="ff-mono text-[11px] px-2 py-0.5 rounded bg-amber-50 text-amber-800 border border-amber-200 shrink-0"
          title="Atraso da entrega em relação à previsão"
        >
          +{lateDays}d
        </span>
      )}
    </div>
  );
}

/* ---------- Offender card ---------- */

function OffenderCard({ rank, offender, totalOrders }) {
  const issues = CATEGORIES
    .filter((c) => c.isProblem)
    .map((c) => ({ ...c, count: offender.categories[c.key] }))
    .filter((c) => c.count > 0)
    .sort((a, b) => b.count - a.count);

  const headline = issues[0];

  return (
    <div className="avoid-break border border-[#1c1c2e]/20 bg-white rounded-sm overflow-hidden flex flex-col">
      <div className="px-5 pt-5 pb-4 border-b border-[#1c1c2e]/10">
        <div className="flex items-start justify-between gap-3 mb-3">
          <div className="ff-display text-5xl font-light leading-none text-[#b45309]">
            {String(rank).padStart(2, '0')}
          </div>
          <div className="text-right">
            <div className="text-[10px] uppercase tracking-[0.2em] text-[#1c1c2e]/50">Volume</div>
            <div className="ff-mono text-sm">{fmtNum(offender.total)} pedidos</div>
            <div className="ff-mono text-[10px] text-[#1c1c2e]/50">{fmtPct(offender.volumeShare)} do total</div>
          </div>
        </div>
        <div className="ff-display text-xl leading-tight">{offender.name}</div>
      </div>

      <div className="px-5 py-4 bg-[#faf7f2]/60 border-b border-[#1c1c2e]/10">
        <div className="text-[10px] uppercase tracking-[0.2em] text-[#1c1c2e]/50 mb-1">Por que é ofensora</div>
        <div className="text-sm leading-snug">
          Concentrou <strong className="ff-mono">{fmtNum(offender.problems)}</strong> ocorrência{offender.problems > 1 ? 's' : ''} —{' '}
          <strong>{fmtPct(offender.problemRate)}</strong> dos pedidos da própria transportadora,
          o que representa <strong className="text-[#b45309]">{fmtPct(offender.shareOfTotal)}</strong> do total geral de {fmtNum(totalOrders)} pedidos.
          {headline && (
            <> Principal problema: <strong>{headline.label.toLowerCase()}</strong> ({fmtNum(headline.count)}).</>
          )}
        </div>
      </div>

      <div className="px-5 py-4 flex-1">
        <div className="text-[10px] uppercase tracking-[0.2em] text-[#1c1c2e]/50 mb-3">Detalhamento das ocorrências</div>
        <ul className="space-y-2">
          {issues.map((c) => {
            const Icon = c.icon;
            const cl = COLOR[c.color];
            return (
              <li key={c.key} className="flex items-center gap-3">
                <div className={`w-7 h-7 rounded-full ${cl.solid} text-white flex items-center justify-center shrink-0`}>
                  <Icon className="w-3.5 h-3.5" strokeWidth={2} />
                </div>
                <div className="flex-1 text-sm">{c.label}</div>
                <div className="ff-mono text-sm">
                  {fmtNum(c.count)}
                  <span className="text-[#1c1c2e]/50 text-xs ml-1.5">
                    ({fmtPct(c.count / offender.total)})
                  </span>
                </div>
              </li>
            );
          })}
        </ul>
      </div>
    </div>
  );
}
