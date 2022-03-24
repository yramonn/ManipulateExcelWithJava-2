package com.sovos;

public class SelectNotaFiscal {
    int id_notaFiscal;
    int serie;
    String nome;
    String descrição;
    int  qtd;
    double valor;
    double ValorTotal;

    public int getId_notaFiscal() {
        return id_notaFiscal;
    }

    public void setId_notaFiscal(int id_notaFiscal) {
        this.id_notaFiscal = id_notaFiscal;
    }

    public int getSerie() {
        return serie;
    }

    public void setSerie(int serie) {
        this.serie = serie;
    }

    public String getNome() {
        return nome;
    }

    public void setNome(String nome) {
        this.nome = nome;
    }

    public String getDescrição() {
        return descrição;
    }

    public void setDescrição(String descricao) {
        this.descrição = descricao;
    }

    public int getQtd() {
        return qtd;
    }

    public void setQtd(int qtd) {
        this.qtd = qtd;
    }

    public double getValor() {
        return valor;
    }

    public void setValor(double valor) {
        this.valor = valor;
    }

    public double getValorTotal() {
        return ValorTotal;
    }

    public void setValorTotal(double valorTotal) {
        ValorTotal = valorTotal;
    }


    public int getId_cliente() {
        return 0;
    }
}
