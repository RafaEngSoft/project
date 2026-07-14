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

    // Busca robusta de metadados antes da tabela começar (limite de linhas)
    const limiteLinhaMetadata = inicioTabelaLinha !== -1 ? inicioTabelaLinha : 15;

    const buscarMetadado = (termos: string[]): number | null => {
      for (const chave of chaves) {
        if (chave.startsWith('!')) continue;
        const celula = planilha[chave];
        if (!celula || celula.v === undefined || celula.v === null) continue;

        const texto = String(celula.v).trim().toLowerCase();
        const corresponde = termos.some(termo => texto.includes(termo));

        if (corresponde) {
          const posicao = XLSX.utils.decode_cell(chave);
          if (posicao.r >= limiteLinhaMetadata) continue;

          // 1. Tentar extrair do próprio texto da célula (ex: "Valor do Contrato: R$ 120.000,00" ou "Avanço da Obra: 75%")
          const numProprio = this.extrairNumeroDoTexto(String(celula.v));

          // Listar células candidatas ao redor (direita, duas à direita, baixo, diagonal)
          const candidatas = [
            { r: posicao.r, c: posicao.c + 1 }, // Direita
            { r: posicao.r, c: posicao.c + 2 }, // Duas à direita
            { r: posicao.r + 1, c: posicao.c }, // Baixo
            { r: posicao.r + 1, c: posicao.c + 1 } // Diagonal
          ];

          for (const cand of candidatas) {
            const ref = XLSX.utils.encode_cell(cand);
            const celCand = planilha[ref];
            if (celCand && celCand.v !== undefined && celCand.v !== null) {
              const num = this.converterNumeroOuNulo(celCand.v);
              if (num !== null && num !== 0) {
                return num;
              }
            }
          }

          // Se não encontrou nas vizinhas, tenta extrair do próprio texto
          if (numProprio !== null) {
            const ehContrato = termos.includes('contrato');
            // Se for contrato, só aceita se for grande ou contiver símbolo monetário, para evitar falsos positivos como ID de contrato
            if (!ehContrato || String(celula.v).includes('R$') || String(celula.v).includes('$') || numProprio > 100) {
              return numProprio;
            }
          }
        }
      }
      return null;
    };

    valorContrato = buscarMetadado(['contrato', 'valor_contrato', 'valor contratado']) || 0;
    mesesCorridos = buscarMetadado(['meses', 'decorrid', 'corridos', 'prazo']) || 0;

    let avancoLido = buscarMetadado(['avanço', 'avanco', 'progresso', 'físico', 'fisico']);
    if (avancoLido !== null) {
      if (avancoLido > 1.0) {
        avancoObra = avancoLido / 100.0;
      } else {
        avancoObra = avancoLido;
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
          let rowHasData = false;
          for (let c = 0; c < 10; c++) {
            if (planilha[XLSX.utils.encode_cell({ r, c })]?.v !== undefined) {
              rowHasData = true;
              break;
            }
          }
          if (!rowHasData) break;
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
                if (!k.includes('desvio') && !k.includes('variação') && !k.includes('variacao') && !k.includes('diferença') && !k.includes('diferenca')) {
                  if (val !== undefined && val !== null) return val;
                }
              }
            }
            if (headerName === 'Valor_Previsto') {
              if ((k.includes('previsto') || k.includes('planejado') || k.includes('orçado') || k.includes('orcado')) && !k.includes('contrato') && !k.includes('%')) {
                if (!k.includes('desvio') && !k.includes('variação') && !k.includes('variacao') && !k.includes('diferença') && !k.includes('diferenca')) {
                  if (val !== undefined && val !== null) return val;
                }
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
                if (!k.includes('desvio') && !k.includes('variação') && !k.includes('variacao') && !k.includes('diferença') && !k.includes('diferenca')) {
                  if (val !== undefined && val !== null) return val;
                }
              }
            }
            if (headerName === 'Previsto_Porc_Contrato') {
              if ((k.includes('previsto') || k.includes('planejado') || k.includes('orçado') || k.includes('orcado')) && (k.includes('contrato') || k.includes('%'))) {
                if (!k.includes('desvio') && !k.includes('variação') && !k.includes('variacao') && !k.includes('diferença') && !k.includes('diferenca')) {
                  if (val !== undefined && val !== null) return val;
                }
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
        let valorRealizado = this.converterNumeroOuNulo(obterValorPorCabecalho('Valor_Realizado'));
        let valorPrevisto = this.converterNumeroOuNulo(obterValorPorCabecalho('Valor_Previsto'));

        let desvioValor = this.converterNumeroOuNulo(obterValorPorCabecalho('Desvio_Valor'));
        let desvioVSPrevisto = this.converterNumeroOuNulo(obterValorPorCabecalho('Desvio_vs_Previsto'));
        let realizadoPorcContrato = this.converterNumeroOuNulo(obterValorPorCabecalho('Realizado_Porc_Contrato'));
        let previstoPorcContrato = this.converterNumeroOuNulo(obterValorPorCabecalho('Previsto_Porc_Contrato'));
        let desvioPorcContrato = this.converterNumeroOuNulo(obterValorPorCabecalho('Desvio_Porc_Contrato'));

        if (mes || valorRealizado !== null || valorPrevisto !== null) {
          const divisorContrato = valorContrato || 1;

          // Normalizar porcentagens se vieram como inteiros do Excel (ex: 8.33 ao invés de 0.0833)
          if (previstoPorcContrato !== null && previstoPorcContrato > 1.0) {
            previstoPorcContrato = previstoPorcContrato / 100.0;
          }
          if (realizadoPorcContrato !== null && realizadoPorcContrato > 1.0) {
            realizadoPorcContrato = realizadoPorcContrato / 100.0;
          }
          if (desvioPorcContrato !== null && Math.abs(desvioPorcContrato) > 1.0) {
            desvioPorcContrato = desvioPorcContrato / 100.0;
          }

          // Preenchimento recíproco de valores baseado nas porcentagens do contrato
          if (valorPrevisto === null && previstoPorcContrato !== null && valorContrato > 0) {
            valorPrevisto = previstoPorcContrato * valorContrato;
          }
          if (valorRealizado === null && realizadoPorcContrato !== null && valorContrato > 0) {
            valorRealizado = realizadoPorcContrato * valorContrato;
          }

          // Preenchimento de porcentagens baseado nos valores se estiverem vazios
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

  extrairNumeroDoTexto(texto: string): number | null {
    if (!texto) return null;
    const match = texto.replace(/R\$\s*/gi, '').trim().match(/[-+]?\s*\d+([\d.,]*)?/);
    if (match) {
      return this.converterNumeroOuNulo(match[0]);
    }
    return null;
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
