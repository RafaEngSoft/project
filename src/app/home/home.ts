import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-home',
  imports: [CommonModule, FormsModule],
  templateUrl: './home.html',
  styleUrl: './home.css',
})
export class homeComponent {
  valorContrato = 0;
  mesesCorridos = 0;
  avancoObra = 0;

  dadosExcel: any[] = [];

  totalRealizado = 0;
  totalPrevisto = 0;
  totalDesvio = 0;
  totalDesvioPercentual = 0;
  totalRealizadoPorcContrato = 0;
  totalPrevistoPorcContrato = 0;
  totalDesvioPorcContrato = 0;

  mesDiarioSelecionado = '';
  textoNotaDiario = '';

  notasDiario: { [mes: string]: string } = {};

  caminhoSvgPrevisto = '';
  caminhoSvgRealizado = '';
  pontosSvgPrevisto: { x: number; y: number; label: string; val: number }[] = [];
  pontosSvgRealizado: { x: number; y: number; label: string; val: number }[] = [];
  linhasGradeY: { y: number; label: string }[] = [];
  rotulosGradeX: { x: number; label: string }[] = [];

  constructor() {
    this.calcularTotais();
    this.calcularDadosAcumulados();
    this.alterarMesDiario();
  }

  lerExcel(event: any) {
    const arquivo = event.target.files[0];
    if (!arquivo) return;

    const leitorArquivo = new FileReader();
    leitorArquivo.readAsArrayBuffer(arquivo);

    leitorArquivo.onload = () => {
      const pastaTrabalho = XLSX.read(leitorArquivo.result, { type: 'array' });
      const nomesPlanilhas = pastaTrabalho.SheetNames;
      const planilha = pastaTrabalho.Sheets[nomesPlanilhas[0]];

      const resultado = this.analisarPlanilhaExcel(planilha);

      if (resultado.linhas && resultado.linhas.length > 0) {
        this.valorContrato = resultado.valorContrato || this.valorContrato;
        this.mesesCorridos = resultado.mesesCorridos || this.mesesCorridos;
        this.avancoObra = resultado.avancoObra || this.avancoObra;
        this.dadosExcel = resultado.linhas;

        if (this.dadosExcel.length > 0) {
          const primeiroMes = this.dadosExcel[0].mes;
          if (primeiroMes && !this.dadosExcel.find(r => r.mes === this.mesDiarioSelecionado)) {
            this.mesDiarioSelecionado = primeiroMes;
          }
        }

        this.calcularTotais();
        this.calcularDadosAcumulados();
        this.alterarMesDiario();
      }
    };
  }

  alterarMesDiario() {
    this.textoNotaDiario = this.notasDiario[this.mesDiarioSelecionado] || '';
  }

  salvarNotaDiario() {
    if (this.mesDiarioSelecionado) {
      this.notasDiario[this.mesDiarioSelecionado] = this.textoNotaDiario;
    }
  }

  calcularTotais() {
    this.totalRealizado = 0;
    this.totalPrevisto = 0;

    if (!this.dadosExcel || this.dadosExcel.length === 0) {
      this.totalDesvio = 0;
      this.totalDesvioPercentual = 0;
      this.totalRealizadoPorcContrato = 0;
      this.totalPrevistoPorcContrato = 0;
      this.totalDesvioPorcContrato = 0;
      return;
    }

    for (const linha of this.dadosExcel) {
      this.totalRealizado += linha.valorRealizado || 0;
      this.totalPrevisto += linha.valorPrevisto || 0;
    }

    this.totalDesvio = this.totalRealizado - this.totalPrevisto;
    this.totalDesvioPercentual = this.totalPrevisto > 0 ? (this.totalDesvio / this.totalPrevisto) : 0;

    const contrato = this.valorContrato || 1;
    this.totalRealizadoPorcContrato = this.totalRealizado / contrato;
    this.totalPrevistoPorcContrato = this.totalPrevisto / contrato;
    this.totalDesvioPorcContrato = this.totalDesvio / contrato;
  }

  calcularDadosAcumulados() {
    if (!this.dadosExcel || this.dadosExcel.length === 0) {
      this.pontosSvgPrevisto = [];
      this.pontosSvgRealizado = [];
      this.caminhoSvgPrevisto = '';
      this.caminhoSvgRealizado = '';
      this.linhasGradeY = [];
      this.rotulosGradeX = [];
      return;
    }

    let previstoAcumulado = 0;
    let realizadoAcumulado = 0;

    this.pontosSvgPrevisto = [];
    this.pontosSvgRealizado = [];

    let ultimoIndiceRealizado = -1;
    for (let i = 0; i < this.dadosExcel.length; i++) {
      if (this.dadosExcel[i].valorRealizado !== null && this.dadosExcel[i].valorRealizado !== undefined) {
        ultimoIndiceRealizado = i;
      }
    }

    const valoresPrevistos: number[] = [];
    const valoresRealizados: number[] = [];

    for (let i = 0; i < this.dadosExcel.length; i++) {
      const linha = this.dadosExcel[i];
      
      if (linha.valorPrevisto !== null && linha.valorPrevisto !== undefined) {
        previstoAcumulado += linha.valorPrevisto;
        valoresPrevistos.push(previstoAcumulado);
      } else {
        valoresPrevistos.push(previstoAcumulado);
      }

      if (i <= ultimoIndiceRealizado && linha.valorRealizado !== null && linha.valorRealizado !== undefined) {
        realizadoAcumulado += linha.valorRealizado;
        valoresRealizados.push(realizadoAcumulado);
      }
    }

    const valorMaximo = Math.max(
      valoresPrevistos[valoresPrevistos.length - 1] || 0,
      valoresRealizados[valoresRealizados.length - 1] || 0,
      100
    );

    const largura = 600;
    const altura = 240;
    const recuoEsquerda = 85;
    const recuoDireita = 30;
    const recuoTopo = 20;
    const recuoBase = 30;

    const larguraDesenho = largura - recuoEsquerda - recuoDireita;
    const alturaDesenho = altura - recuoTopo - recuoBase;

    const totalItens = this.dadosExcel.length;

    this.pontosSvgPrevisto = valoresPrevistos.map((val, i) => {
      const x = recuoEsquerda + (totalItens > 1 ? (i / (totalItens - 1)) * larguraDesenho : 0);
      const y = (recuoTopo + alturaDesenho) - (valorMaximo > 0 ? (val / valorMaximo) : 0) * alturaDesenho;
      return { x, y, label: this.dadosExcel[i].mes, val };
    });

    this.pontosSvgRealizado = valoresRealizados.map((val, i) => {
      const x = recuoEsquerda + (totalItens > 1 ? (i / (totalItens - 1)) * larguraDesenho : 0);
      const y = (recuoTopo + alturaDesenho) - (valorMaximo > 0 ? (val / valorMaximo) : 0) * alturaDesenho;
      return { x, y, label: this.dadosExcel[i].mes, val };
    });

    if (this.pontosSvgPrevisto.length > 0) {
      this.caminhoSvgPrevisto = 'M ' + this.pontosSvgPrevisto.map(p => `${p.x} ${p.y}`).join(' L ');
    } else {
      this.caminhoSvgPrevisto = '';
    }

    if (this.pontosSvgRealizado.length > 0) {
      this.caminhoSvgRealizado = 'M ' + this.pontosSvgRealizado.map(p => `${p.x} ${p.y}`).join(' L ');
    } else {
      this.caminhoSvgRealizado = '';
    }

    this.linhasGradeY = [];
    for (let j = 0; j <= 4; j++) {
      const proporcao = j / 4;
      const valor = proporcao * valorMaximo;
      const y = (recuoTopo + alturaDesenho) - proporcao * alturaDesenho;
      this.linhasGradeY.push({
        y,
        label: this.formatarMoeda(valor),
      });
    }

    this.rotulosGradeX = this.pontosSvgPrevisto.map(p => ({
      x: p.x,
      label: p.label ? p.label.substring(0, 3) : '',
    }));
  }

  formatarMoeda(valor: number): string {
    return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL', maximumFractionDigits: 0 }).format(valor);
  }

  analisarPlanilhaExcel(planilha: XLSX.WorkSheet) {
    const chaves = Object.keys(planilha);

    let valorContrato = 0;
    let mesesCorridos = 0;
    let avancoObra = 0;

    for (const chave of chaves) {
      if (chave.startsWith('!')) continue;
      const celula = planilha[chave];
      if (!celula || celula.v === undefined || celula.v === null) continue;

      const valorTexto = String(celula.v).trim().toLowerCase();
      const posicao = XLSX.utils.decode_cell(chave);

      // Limit metadata check to the top rows (rows 1-5, which have index 0-4)
      if (posicao.r < 5) {
        if (valorTexto.includes('contrato') || valorTexto.includes('valor_contrato')) {
          const proximaCelula = XLSX.utils.encode_cell({ r: posicao.r, c: posicao.c + 1 });
          const val = planilha[proximaCelula]?.v;
          if (val !== undefined && val !== null) {
            valorContrato = this.converterNumeroOuNulo(val) || 0;
          }
        }

        if (valorTexto.includes('meses') || valorTexto.includes('decorrid')) {
          const proximaCelula = XLSX.utils.encode_cell({ r: posicao.r, c: posicao.c + 1 });
          const val = planilha[proximaCelula]?.v;
          if (val !== undefined && val !== null) {
            mesesCorridos = this.converterNumeroOuNulo(val) || 0;
          }
        }

        if (valorTexto.includes('avanço') || valorTexto.includes('avanco') || valorTexto.includes('progresso')) {
          const proximaCelula = XLSX.utils.encode_cell({ r: posicao.r, c: posicao.c + 1 });
          const val = planilha[proximaCelula]?.v;
          if (val !== undefined && val !== null) {
            avancoObra = this.converterNumeroOuNulo(val) || 0;
          }
        }
      }
    }

    let inicioTabelaLinha = -1;
    const indicesColunas: { [chave: string]: number } = {};

    for (let r = 0; r < 30; r++) {
      let ehLinhaCabecalho = false;
      for (let c = 0; c < 15; c++) {
        const chaveCelula = XLSX.utils.encode_cell({ r, c });
        const celula = planilha[chaveCelula];
        if (celula && celula.v !== undefined && celula.v !== null) {
          const val = String(celula.v).trim().toLowerCase();
          if (val === 'mês' || val === 'mes' || val === 'valor_realizado' || val.includes('previsto') || val.includes('realizado')) {
            ehLinhaCabecalho = true;
            break;
          }
        }
      }

      if (ehLinhaCabecalho) {
        inicioTabelaLinha = r;
        for (let c = 0; c < 15; c++) {
          const chaveCelula = XLSX.utils.encode_cell({ r, c });
          const celula = planilha[chaveCelula];
          if (celula && celula.v !== undefined && celula.v !== null) {
            const val = String(celula.v).trim();
            indicesColunas[val] = c;
          }
        }
        break;
      }
    }

    const linhas: any[] = [];
    if (inicioTabelaLinha !== -1) {
      for (let r = inicioTabelaLinha + 1; r < inicioTabelaLinha + 25; r++) {
        let colMes = -1;
        for (const key of Object.keys(indicesColunas)) {
          if (key.toLowerCase() === 'mês' || key.toLowerCase() === 'mes') {
            colMes = indicesColunas[key];
            break;
          }
        }
        if (colMes === -1) colMes = 0;

        const chaveMes = XLSX.utils.encode_cell({ r, c: colMes });
        const valorMes = planilha[chaveMes]?.v;

        if (valorMes === undefined || valorMes === null || String(valorMes).trim() === '') {
          let linhaTemDados = false;
          for (let c = 0; c < 10; c++) {
            if (planilha[XLSX.utils.encode_cell({ r, c })]?.v !== undefined) {
              linhaTemDados = true;
              break;
            }
          }
          if (!linhaTemDados) break;
        }

        if (valorMes && (String(valorMes).toLowerCase().includes('total') || String(valorMes).toLowerCase().includes('geral'))) {
          continue;
        }

        const obterValorPorCabecalho = (headerName: string) => {
          for (const key of Object.keys(indicesColunas)) {
            const k = key.toLowerCase();
            const col = indicesColunas[key];
            const val = planilha[XLSX.utils.encode_cell({ r, c: col })]?.v;

            if (headerName === 'Valor_Realizado') {
              if ((k.includes('realizado') || k.includes('medido')) && !k.includes('contrato') && !k.includes('%')) {
                if (val !== undefined && val !== null) return val;
              }
            }
            if (headerName === 'Valor_Previsto') {
              if ((k.includes('previsto') || k.includes('planejado') || k.includes('orçado') || k.includes('orcado')) && !k.includes('contrato') && !k.includes('%')) {
                if (val !== undefined && val !== null) return val;
              }
            }
            if (headerName === 'Desvio_Valor') {
              if ((k.includes('desvio') || k.includes('diferença') || k.includes('diferenca')) && !k.includes('contrato') && !k.includes('%')) {
                if (val !== undefined && val !== null) return val;
              }
            }
            if (headerName === 'Desvio_vs_Previsto') {
              if ((k.includes('desvio') || k.includes('variação') || k.includes('variacao')) && k.includes('previsto') && !k.includes('contrato')) {
                if (val !== undefined && val !== null) return val;
              }
              if ((k.includes('desvio') || k.includes('variação') || k.includes('variacao')) && k.includes('%') && !k.includes('contrato')) {
                if (val !== undefined && val !== null) return val;
              }
            }
            if (headerName === 'Realizado_Porc_Contrato') {
              if ((k.includes('realizado') || k.includes('medido')) && (k.includes('contrato') || k.includes('%'))) {
                if (val !== undefined && val !== null) return val;
              }
            }
            if (headerName === 'Previsto_Porc_Contrato') {
              if ((k.includes('previsto') || k.includes('planejado') || k.includes('orçado') || k.includes('orcado')) && (k.includes('contrato') || k.includes('%'))) {
                if (val !== undefined && val !== null) return val;
              }
            }
            if (headerName === 'Desvio_Porc_Contrato') {
              if ((k.includes('desvio') || k.includes('diferença') || k.includes('diferenca')) && (k.includes('contrato') || k.includes('%'))) {
                if (!k.includes('previsto') && !k.includes('prev')) {
                  if (val !== undefined && val !== null) return val;
                }
              }
            }
          }
          return null;
        };

        const mes = valorMes ? String(valorMes).trim() : '';
        const valorRealizado = this.converterNumeroOuNulo(obterValorPorCabecalho('Valor_Realizado'));
        const valorPrevisto = this.converterNumeroOuNulo(obterValorPorCabecalho('Valor_Previsto'));
        
        let desvioValor = this.converterNumeroOuNulo(obterValorPorCabecalho('Desvio_Valor'));
        let desvioVSPrevisto = this.converterNumeroOuNulo(obterValorPorCabecalho('Desvio_vs_Previsto'));
        let realizadoPorcContrato = this.converterNumeroOuNulo(obterValorPorCabecalho('Realizado_Porc_Contrato'));
        let previstoPorcContrato = this.converterNumeroOuNulo(obterValorPorCabecalho('Previsto_Porc_Contrato'));
        let desvioPorcContrato = this.converterNumeroOuNulo(obterValorPorCabecalho('Desvio_Porc_Contrato'));

        if (mes || valorRealizado !== null || valorPrevisto !== null) {
          const divisorContrato = valorContrato || 1;
          
          if (valorPrevisto !== null && previstoPorcContrato === null) {
            previstoPorcContrato = valorPrevisto / divisorContrato;
          }

          if (valorRealizado !== null) {
            if (realizadoPorcContrato === null) {
              realizadoPorcContrato = valorRealizado / divisorContrato;
            }
            if (valorPrevisto !== null) {
              if (desvioValor === null) {
                desvioValor = valorRealizado - valorPrevisto;
              }
              if (desvioVSPrevisto === null) {
                desvioVSPrevisto = valorPrevisto !== 0 ? (desvioValor / valorPrevisto) : 0;
              }
              if (desvioPorcContrato === null) {
                desvioPorcContrato = realizadoPorcContrato - (previstoPorcContrato || 0);
              }
            }
          }

          linhas.push({
            mes,
            valorRealizado,
            valorPrevisto,
            desvioValor,
            desvioVSPrevisto,
            realizadoPorcContrato,
            previstoPorcContrato,
            desvioPorcContrato,
          });
        }
      }
    }

    return {
      valorContrato,
      mesesCorridos,
      avancoObra,
      linhas,
    };
  }

  converterNumeroOuNulo(val: any): number | null {
    if (val === null || val === undefined || String(val).trim() === '') return null;
    if (typeof val === 'number') return val;

    let str = String(val).replace(/R\$\s*/g, '').trim();
    if (str === '-' || str === '') return null;

    if (str.includes('.') && str.includes(',')) {
      str = str.replace(/\./g, '').replace(/,/g, '.');
    } else if (str.includes(',')) {
      str = str.replace(/,/g, '.');
    }

    let ehPercentual = false;
    if (str.includes('%')) {
      str = str.replace(/%/g, '');
      ehPercentual = true;
    }

    if (str.startsWith('(') && str.endsWith(')')) {
      str = '-' + str.substring(1, str.length - 1);
    }

    let num = parseFloat(str);
    if (isNaN(num)) return null;

    if (ehPercentual) {
      num = num / 100.0;
    }
    return num;
  }
}
