package com.sovos;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.LinkedList;
import java.util.List;

public class Main {
    static final String fileName = "C://Users//Ramon.Silva//Downloads//excel//NotaFiscal.xlsx";

    public static void main(String[] args) throws IOException {

        Statement stmt = null;
        ResultSet rs = null;
        Connection conn = null;

        List<SelectNotaFiscal> listSelectNotaFiscal = new LinkedList<SelectNotaFiscal>();

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheetSelectNotaFiscal1 = workbook.createSheet("NotaFiscal1");
        XSSFSheet sheetSelectNotaFiscal2 = workbook.createSheet("NotaFiscal2");
        XSSFSheet sheetSelectNotaFiscal3 = workbook.createSheet("NotaFiscal3");
        XSSFSheet sheetSelectNotaFiscal4 = workbook.createSheet("NotaFiscal4");
        XSSFSheet sheetSelectNotaFiscal5 = workbook.createSheet("NotaFiscal5");


        int rownum = 0;

        try {
            Conexao c = new Conexao();
            conn = c.getNewConecction();
            stmt = conn.createStatement();


            if (stmt.execute("SELECT nf.id_notaFiscal,nf.serie, cl.nome,pd.descrição, itnf.qtd,pd.valor, (itnf.qtd * pd.valor) as ValorTotal\n" +
                    " FROM nota_fiscal nf, cliente cl, produto pd, itens_notafiscal itnf\n" +
                    "WHERE cl.id_cliente = nf.cliente and  nf.id_notaFiscal = itnf.id_notaFiscal and pd.id_produtos = itnf.id_produtos;")) {
                rs = stmt.getResultSet();
            }
            int id = 0;
            int newid = 0;
            int totalid = 0;
            while (rs.next()) {
                SelectNotaFiscal snf = new SelectNotaFiscal();
                snf.setId_notaFiscal(rs.getInt("id_notaFiscal"));
                snf.setNome(rs.getString("nome"));
                snf.setSerie(rs.getInt("serie"));
                snf.setDescrição(rs.getString("descrição"));
                snf.setQtd(rs.getInt("qtd"));
                snf.setValor(rs.getDouble("valor"));
                snf.setValorTotal(rs.getDouble("ValorTotal"));
                newid = snf.id_notaFiscal;
                if (id != newid) {
                    totalid++;
                    id = newid;
                }
                listSelectNotaFiscal.add(snf);
            }

            for (int i = 0; i < totalid; i++) {
                if (workbook.getNumberOfSheets() < totalid) {
                    workbook.createSheet("NF" + i);
                }
            }


            int index = 0;
            while (index < totalid) {

                Row rowCab = workbook.getSheetAt(index).createRow(rownum++);
                int cellnumCab = 0;
                int totalNF=0;

                Cell cellId_notaFiscal = rowCab.createCell(cellnumCab++);
                Cell cellSerie = rowCab.createCell(cellnumCab++);
                Cell cellNome = rowCab.createCell(cellnumCab++);
                Cell cellDescrição = rowCab.createCell(cellnumCab++);
                Cell cellQtd = rowCab.createCell(cellnumCab++);
                Cell cellValor = rowCab.createCell(cellnumCab++);
                Cell cellValorTotal = rowCab.createCell(cellnumCab++);
                Cell cellValorNF = rowCab.createCell(cellnumCab++);



                cellId_notaFiscal.setCellValue("id_notaFiscal");
                cellSerie.setCellValue("serie");
                cellNome.setCellValue("nome");
                cellDescrição.setCellValue("descricao ");
                cellQtd.setCellValue("qtd");
                cellValor.setCellValue("valor");
                cellValorTotal.setCellValue("ValorTotal");

                cellValorNF.setCellValue("ValorTotalNF");



                for (SelectNotaFiscal snf : listSelectNotaFiscal) {
                    if (snf.getId_notaFiscal() == index + 1) {

                        Row row = workbook.getSheetAt(index).createRow(rownum++);
                        int cellnum = 0;

                        cellId_notaFiscal = row.createCell(cellnum++);
                        cellSerie = row.createCell(cellnum++);
                        cellNome = row.createCell(cellnum++);
                        cellDescrição = row.createCell(cellnum++);
                        cellQtd = row.createCell(cellnum++);
                        cellValor = row.createCell(cellnum++);
                        cellValorTotal = row.createCell(cellnum++);

                        cellId_notaFiscal.setCellValue(snf.getId_notaFiscal());
                        cellValor.setCellValue(snf.getId_cliente());
                        cellNome.setCellValue(snf.getNome());
                        cellSerie.setCellValue(snf.getSerie());
                        cellDescrição.setCellValue(snf.getDescrição());
                        cellQtd.setCellValue(snf.getQtd());
                        cellValor.setCellValue(snf.getValor());
                        cellValorTotal.setCellValue(snf.getValorTotal());
                        totalNF += snf.getValorTotal();
                    }

                }
                cellValorTotal = rowCab.createCell(cellnumCab++);
                cellValorTotal.setCellValue(totalNF);
                index++;
                rownum = 0;

            }

            FileOutputStream out = new FileOutputStream(new File(Main.fileName));
            workbook.write(out);
            out.close();

        } catch (SQLException ex) {
            System.out.println("SQLException: " + ex.getMessage());
            System.out.println("SQLState: " + ex.getSQLState());
            System.out.println("VendorError: " + ex.getErrorCode());
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            System.out.println("Arquivo Excel não encontrado!");
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Erro na edição do arquivo");
        } finally {
            if (rs != null) {
                try {
                    rs.close();
                } catch (SQLException sqlEx) {

                }
            }

            if (stmt != null) {
                try {
                    stmt.close();
                } catch (SQLException sqlEx) {

                }
            }

            if (conn != null) {
                try {
                    conn.close();
                } catch (SQLException sqlEx) {

                }
            }
        }
    }
}
