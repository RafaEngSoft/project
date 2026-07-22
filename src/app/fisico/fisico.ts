import { Component, ChangeDetectorRef, NgZone } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import * as XLSX from 'xlsx';
import { ModalService } from '../modal.service';

@Component({
  selector: 'app-fisico',
  imports: [CommonModule, FormsModule],
  templateUrl: './fisico.html',
  styleUrl: './fisico.css',
})
export class Fisico {
  // Acompanhamento de Prazo — Curva S
  dataInicioObra: any = null;
  dataStatusMedicao: any = null;
  duracaoLinhaBase: number | null = null;
  duracaoRealAtualizada: number | null = null;
  tempoTranscorrido: number | null = null;
  avancoFisicoEsperado: number | null = null;
  dataTerminoLinhaBase: any = null;
  dataTerminoAtual: any = null;
  variacaoTermino: number | null = null;

  // Avanço Físico Real e Desvio
  avancoFisicoRealAcumulado: number | null = null;
  desvioAvanco: number | null = null;

  // Tabela Mensal
  dadosExcel: any[] = [];

  // Notas / Diário
  mesDiarioSelecionado = '';
  textoNotaDiario = '';
  notaSalvaMensagem = false;
  notasDiario: { [mes: string]: string } = {};

  // Curva S (SVG)
  caminhoSvgPrevisto = '';
  caminhoSvgRealizado = '';
  pontosSvgPrevisto: { x: number; y: number; label: string; val: number }[] = [];
  pontosSvgRealizado: { x: number; y: number; label: string; val: number }[] = [];
  linhasGradeY: { y: number; label: string }[] = [];
  rotulosGradeX: { x: number; label: string }[] = [];

  constructor(
    private ngZone: NgZone,
    private cdr: ChangeDetectorRef,
    public modalService: ModalService
  ) {
    this.calcularDadosAcumulados();
    this.alterarMesDiario();
  }

  abrirCurvaS() {
    this.modalService.showSCurveModal.set(true);
  }

  closeSCurveModal() {
    this.modalService.showSCurveModal.set(false);
  }

  lerExcel(event: any) {
    console.log('1. lerExcel (Fisico) foi acionado!', event.target.files);

    const arquivo = event.target.files[0];
    if (!arquivo) {
      console.log('Erro: Nenhum arquivo encontrado no input.');
      return;
    }

    const leitorArquivo = new FileReader();

    leitorArquivo.onload = () => {
      console.log('2. FileReader.onload terminou de ler o arquivo do disco.');

      this.ngZone.run(() => {
        const pastaTrabalho = XLSX.read(leitorArquivo.result, { type: 'array' });
        const nomesPlanilhas = pastaTrabalho.SheetNames;
        const planilha = pastaTrabalho.Sheets[nomesPlanilhas[0]];

        const resultado = this.analisarPlanilhaFisico(planilha);
        console.log('3. Resultado do processamento da planilha:', resultado);

        if (resultado.linhas && resultado.linhas.length > 0) {
          this.dataInicioObra = resultado.dataInicioObra;
          this.dataStatusMedicao = resultado.dataStatusMedicao;
          this.duracaoLinhaBase = resultado.duracaoLinhaBase;
          this.duracaoRealAtualizada = resultado.duracaoRealAtualizada;
          this.tempoTranscorrido = resultado.tempoTranscorrido;
          this.avancoFisicoEsperado = resultado.avancoFisicoEsperado;
          this.dataTerminoLinhaBase = resultado.dataTerminoLinhaBase;
          this.dataTerminoAtual = resultado.dataTerminoAtual;
          this.variacaoTermino = resultado.variacaoTermino;
          this.avancoFisicoRealAcumulado = resultado.avancoFisicoRealAcumulado;
          this.desvioAvanco = resultado.desvioAvanco;

          this.dadosExcel = resultado.linhas;

          if (this.dadosExcel.length > 0) {
            const primeiroMes = this.dadosExcel[0].mes;
            if (primeiroMes && !this.dadosExcel.find(r => r.mes === this.mesDiarioSelecionado)) {
              this.mesDiarioSelecionado = primeiroMes;
            }
          }

          this.calcularDadosAcumulados();
          this.alterarMesDiario();
          console.log('4. Estados atualizados e cálculos concluídos!');
          this.cdr.detectChanges();
        } else {
          console.log('Erro: Nenhuma linha de dados válida foi extraída da planilha.');
        }
      });
      event.target.value = '';
    };

    leitorArquivo.readAsArrayBuffer(arquivo);
  }

  alterarMesDiario() {
    this.textoNotaDiario = this.notasDiario[this.mesDiarioSelecionado] || '';
  }

  salvarNotaDiario() {
    if (this.mesDiarioSelecionado) {
      this.notasDiario[this.mesDiarioSelecionado] = this.textoNotaDiario;
      this.notaSalvaMensagem = true;
      setTimeout(() => {
        this.notaSalvaMensagem = false;
      }, 3000);
    }
  }

  obterMesesComNotas(): string[] {
    return Object.keys(this.notasDiario).filter(mes => this.notasDiario[mes] && this.notasDiario[mes].trim() !== '');
  }

  selecionarMesDiario(mes: string) {
    this.mesDiarioSelecionado = mes;
    this.alterarMesDiario();
  }

  formatarData(serial: any): string {
    if (serial === null || serial === undefined || isNaN(Number(serial))) {
      if (typeof serial === 'string') return serial;
      return '-';
    }
    const num = Number(serial);
    const date = new Date(1899, 11, 30);
    date.setDate(date.getDate() + num);
    return date.toLocaleDateString('pt-BR');
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

    const valoresPrevistos: number[] = [];
    const valoresRealizados: number[] = [];
    const mesesComRealizado: number[] = [];

    for (let i = 0; i < this.dadosExcel.length; i++) {
      const linha = this.dadosExcel[i];
      const p = linha.avancoPlanejado !== null ? linha.avancoPlanejado : 0;
      valoresPrevistos.push(p);

      if (linha.avancoReal !== null && linha.avancoReal !== undefined) {
        valoresRealizados.push(linha.avancoReal);
        mesesComRealizado.push(i);
      }
    }

    // Y max is 1.0 (100%)
    const valorMaximo = 1.0;

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
      const y = (recuoTopo + alturaDesenho) - (val / valorMaximo) * alturaDesenho;
      return { x, y, label: this.dadosExcel[i].mes, val };
    });

    this.pontosSvgRealizado = valoresRealizados.map((val, idx) => {
      const originalIdx = mesesComRealizado[idx];
      const x = recuoEsquerda + (totalItens > 1 ? (originalIdx / (totalItens - 1)) * larguraDesenho : 0);
      const y = (recuoTopo + alturaDesenho) - (val / valorMaximo) * alturaDesenho;
      return { x, y, label: this.dadosExcel[originalIdx].mes, val };
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
        label: new Intl.NumberFormat('pt-BR', { style: 'percent', minimumFractionDigits: 0 }).format(valor),
      });
    }

    this.rotulosGradeX = this.pontosSvgPrevisto.map(p => ({
      x: p.x,
      label: p.label ? p.label.substring(0, 3) : '',
    }));
  }

  normalizarTexto(t: string): string {
    if (!t) return '';
    return String(t).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  }

  analisarPlanilhaFisico(planilha: XLSX.WorkSheet) {
    const chaves = Object.keys(planilha);

    let dataInicioObra: any = null;
    let dataStatusMedicao: any = null;
    let duracaoLinhaBase: number | null = null;
    let duracaoRealAtualizada: number | null = null;
    let tempoTranscorrido: number | null = null;
    let avancoFisicoEsperado: number | null = null;
    let dataTerminoLinhaBase: any = null;
    let dataTerminoAtual: any = null;
    let variacaoTermino: number | null = null;
    let avancoFisicoRealAcumulado: number | null = null;
    let desvioAvanco: number | null = null;

    const buscarValorAoLado = (termos: string[]): any => {
      const termosNorm = termos.map(t => this.normalizarTexto(t));
      for (const chave of chaves) {
        if (chave.startsWith('!')) continue;
        const celula = planilha[chave];
        if (!celula || celula.v === undefined || celula.v === null) continue;

        const texto = this.normalizarTexto(String(celula.v));
        const corresponde = termosNorm.some(termo => texto === termo || texto.includes(termo));

        if (corresponde) {
          const posicao = XLSX.utils.decode_cell(chave);
          // Verificar vizinhos c+1, c+2, c+3, c+4
          for (let offset = 1; offset <= 4; offset++) {
            const ref = XLSX.utils.encode_cell({ r: posicao.r, c: posicao.c + offset });
            const celCand = planilha[ref];
            if (celCand && celCand.v !== undefined && celCand.v !== null) {
              return celCand.v;
            }
          }
        }
      }
      return null;
    };

    dataInicioObra = buscarValorAoLado(['data de início da obra', 'data de inicio da obra', 'início da obra']);
    dataStatusMedicao = buscarValorAoLado(['data de status (medição)', 'data de status', 'data de medição']);
    duracaoLinhaBase = this.converterNumeroOuNulo(buscarValorAoLado(['duração linha de base (dias)', 'duração linha de base', 'duracao linha de base']));
    duracaoRealAtualizada = this.converterNumeroOuNulo(buscarValorAoLado(['duração real / atualizada (dias)', 'duração real / atualizada', 'duração real', 'duracao real']));
    tempoTranscorrido = this.converterNumeroOuNulo(buscarValorAoLado(['tempo transcorrido (dias)', 'tempo transcorrido']));
    avancoFisicoEsperado = this.converterNumeroOuNulo(buscarValorAoLado(['avanço físico esperado (%)', 'avanço físico esperado', 'avanco fisico esperado']));
    dataTerminoLinhaBase = buscarValorAoLado(['data de término — linha de base', 'data de término linha de base', 'término — linha de base']);
    dataTerminoAtual = buscarValorAoLado(['data de término — atual', 'data de término atual', 'término — atual']);
    variacaoTermino = this.converterNumeroOuNulo(buscarValorAoLado(['variação de término (dias)', 'variação de término', 'variacao de termino']));
    avancoFisicoRealAcumulado = this.converterNumeroOuNulo(buscarValorAoLado(['avanço físico real acumulado (%)', 'avanço físico real acumulado', 'avanco fisico real acumulado']));
    desvioAvanco = this.converterNumeroOuNulo(buscarValorAoLado(['desvio de avanço (real − planejado)', 'desvio de avanço', 'desvio de avanco']));

    // Se as porcentagens vieram como maiores que 1.0 (ex: 55.06 em vez de 0.5506)
    if (avancoFisicoEsperado !== null && avancoFisicoEsperado > 1.0) {
      avancoFisicoEsperado = avancoFisicoEsperado / 100.0;
    }
    if (avancoFisicoRealAcumulado !== null && avancoFisicoRealAcumulado > 1.0) {
      avancoFisicoRealAcumulado = avancoFisicoRealAcumulado / 100.0;
    }
    if (desvioAvanco !== null && Math.abs(desvioAvanco) > 1.0) {
      desvioAvanco = desvioAvanco / 100.0;
    }

    // Achar o início da seção física (para evitar encontrar os cabeçalhos da seção financeira)
    let secaoFisicaLinha = 20; // Fallback
    for (let r = 0; r < 100; r++) {
      let achou = false;
      for (let c = 0; c < 10; c++) {
        const chaveCelula = XLSX.utils.encode_cell({ r, c });
        const celula = planilha[chaveCelula];
        if (celula && celula.v !== undefined && celula.v !== null) {
          const val = this.normalizarTexto(String(celula.v));
          if (val.includes('acompanhamento de prazo') || val.includes('curva s (avanco fisico)') || val.includes('curva s (avanço físico)')) {
            secaoFisicaLinha = r;
            achou = true;
            break;
          }
        }
      }
      if (achou) break;
    }

    // Agora, buscar a tabela Curva S Mensal a partir de secaoFisicaLinha
    let inicioTabelaLinha = -1;
    const indicesColunas: { [chave: string]: number } = {};

    for (let r = secaoFisicaLinha; r < 100; r++) {
      let temMes = false;
      let temAvancoPlanejado = false;
      let temAvancoReal = false;

      for (let c = 0; c < 15; c++) {
        const chaveCelula = XLSX.utils.encode_cell({ r, c });
        const celula = planilha[chaveCelula];
        if (celula && celula.v !== undefined && celula.v !== null) {
          const val = this.normalizarTexto(String(celula.v));
          if (val === 'mes' || val === 'mês') {
            temMes = true;
          }
          if (val.includes('planejado') && val.includes('acumulado')) {
            temAvancoPlanejado = true;
          }
          if (val.includes('real') && val.includes('acumulado')) {
            temAvancoReal = true;
          }
        }
      }

      if (temMes && (temAvancoPlanejado || temAvancoReal)) {
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
          const kNorm = this.normalizarTexto(key);
          if (kNorm === 'mes' || kNorm === 'mês') {
            colMes = indicesColunas[key];
            break;
          }
        }
        if (colMes === -1) colMes = 0;

        const chaveMes = XLSX.utils.encode_cell({ r, c: colMes });
        const valorMes = planilha[chaveMes]?.v;
        const mes = valorMes ? String(valorMes).trim() : '';

        if (!mes ||
            this.normalizarTexto(mes) === 'mes' ||
            this.normalizarTexto(mes).includes('total') ||
            this.normalizarTexto(mes).includes('obs') ||
            this.normalizarTexto(mes).includes('acumulado')) {
          let rowHasData = false;
          for (let c = 0; c < 10; c++) {
            if (planilha[XLSX.utils.encode_cell({ r, c })]?.v !== undefined) {
              rowHasData = true;
              break;
            }
          }
          if (!rowHasData) {
            break;
          }
          continue;
        }

        const obterValorPorCabecalho = (termosHeader: string[]) => {
          const termosNorm = termosHeader.map(t => this.normalizarTexto(t));
          for (const key of Object.keys(indicesColunas)) {
            const k = this.normalizarTexto(key);
            const col = indicesColunas[key];
            const val = planilha[XLSX.utils.encode_cell({ r, c: col })]?.v;

            const matches = termosNorm.every(termo => k.includes(termo));
            if (matches && val !== undefined && val !== null) {
              return val;
            }
          }
          return null;
        };

        let dataRef = obterValorPorCabecalho(['data', 'referência']) || obterValorPorCabecalho(['data', 'referencia']);
        let avancoPlanejado = this.converterNumeroOuNulo(obterValorPorCabecalho(['planejado', 'acumulado']) || obterValorPorCabecalho(['esperado', 'acumulado']));
        let avancoReal = this.converterNumeroOuNulo(obterValorPorCabecalho(['real', 'acumulado']));
        let desvio = this.converterNumeroOuNulo(obterValorPorCabecalho(['desvio', 'avanço']) || obterValorPorCabecalho(['desvio', 'avanco']));

        // Normalizar se vieram como inteiros
        if (avancoPlanejado !== null && avancoPlanejado > 1.0) {
          avancoPlanejado = avancoPlanejado / 100.0;
        }
        if (avancoReal !== null && avancoReal > 1.0) {
          avancoReal = avancoReal / 100.0;
        }
        if (desvio !== null && Math.abs(desvio) > 1.0) {
          desvio = desvio / 100.0;
        }

        linhas.push({
          mes,
          dataRef,
          avancoPlanejado,
          avancoReal,
          desvio
        });
      }
    }

    return {
      dataInicioObra,
      dataStatusMedicao,
      duracaoLinhaBase,
      duracaoRealAtualizada,
      tempoTranscorrido,
      avancoFisicoEsperado,
      dataTerminoLinhaBase,
      dataTerminoAtual,
      variacaoTermino,
      avancoFisicoRealAcumulado,
      desvioAvanco,
      linhas
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
