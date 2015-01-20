/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package my.pruebaxls;

/**
 *
 * @author D.Abrilot
 */

import java.sql.*;
import java.text.NumberFormat.*;

import java.io.*;
import java.util.Calendar;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;


public class PruebaXls {



    /**
     * Crea una hoja Excel y la guarda.
     * author: Daniel
     * @param args
     */

    //metodo que llena el libro excel y retorn "i", para verificar si fueron guardados datos
    public static int llena(HSSFRow fila, int i, ResultSet result, HSSFSheet hoja, HSSFCellStyle style) throws SQLException{
        String[] datos = {"SUCS01","NPAR01","PROC01","DESC01",
                          "STKF01","UNIM01","TRNS01","RNVN01",
                          "RFAT01","STKV01","MINM01","MAXM01",
                          "TPCL01","SLDF01"};  
        while(result.next()){
            // Se crea una fila dentro de la hoja
            fila = hoja.createRow(i);  
            for(int next = 0;next<datos.length;next++){
                HSSFCell celda = fila.createCell((short) next);
                HSSFRichTextString texto = new HSSFRichTextString(result.getString(datos[next]));
                celda.setCellValue(texto);
                celda.setCellStyle(style);
            }
            i++;      
        }
        return i;
    } 
    
    public static void main(String [] args) {
        // Se crea el libro
        HSSFWorkbook libro = new HSSFWorkbook();
        String[] cabecera = {"SUCURSAL","N° DE PARTE","PROCEDENCIA","DESCRIPCIÓN",
                             "STOCK FISICO","UNIDAD DE MEDIDA","TRANSITO",
                             "RESERVA POR NOTA DE VENTA","RESERVA POR FACTURACIÓN ANTICIPADA",
                             "STOCK VIRTUAL","MINIMO","MAXIMO","TIPO CALCULO","SALDO FINAL"};
        //Para la cabecera y las hojas
                                                                 

        //Creo conexion al AS400 por cada archivo se hace una conexion
        Connection conectado401 = null;
        //String database401 = "amsdsgr.phpven";
        String database401 = "atgdsgr.rsldsk01";
        String ipserver = "192.168.2.118";
        String user = "USRCLIENTE";
        String pass = "CLIENTE";

        //obtengo el año actual
        Calendar c1 = Calendar.getInstance();
        int anno = (c1.get(Calendar.YEAR));
        //System.out.println(anno);

        //creo variable que tendra el numero de fila
        int i=1;

        try{
            DriverManager.registerDriver(new com.ibm.as400.access.AS400JDBCDriver());
            String urlconexion401 = "jdbc:as400://" + ipserver + "/" + database401;
            conectado401 = DriverManager.getConnection(urlconexion401, user, pass);

            String sql401 = "select * from " + database401 + " where sucs01 = 1";
            String sql402 = "select * from " + database401 + " where sucs01 = 2";
            String sql403 = "select * from " + database401 + " where sucs01 = 3";
            String sql405 = "select * from " + database401 + " where sucs01 = 5";
            String sql406 = "select * from " + database401 + " where sucs01 = 6";
            String sql407 = "select * from " + database401 + " where sucs01 = 7";

            Statement s401 = conectado401.createStatement();
            Statement s402 = conectado401.createStatement();
            Statement s403 = conectado401.createStatement();
            Statement s405 = conectado401.createStatement();
            Statement s406 = conectado401.createStatement();
            Statement s407 = conectado401.createStatement();

            ResultSet res401 = s401.executeQuery(sql401);
            ResultSet res402 = s402.executeQuery(sql402);
            ResultSet res403 = s403.executeQuery(sql403);
            ResultSet res405 = s405.executeQuery(sql405);
            ResultSet res406 = s406.executeQuery(sql406);
            ResultSet res407 = s407.executeQuery(sql407);
                
            //comienzo la conexion
             try{         
                //por hojas
                for(int j = 0;j<6;j++){
                    HSSFSheet hoja = libro.createSheet();

                    HSSFCellStyle style=libro.createCellStyle();
                    style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                    style.setBorderTop(HSSFCellStyle.BORDER_THIN);
                    style.setBorderRight(HSSFCellStyle.BORDER_THIN);
                    style.setBorderLeft(HSSFCellStyle.BORDER_THIN);

                    HSSFRow fila = hoja.createRow(0);
                    for(int k = 0;k<cabecera.length;k++){

                        HSSFCell celda = fila.createCell((short) k);
                        HSSFRichTextString texto = new HSSFRichTextString(cabecera[k]);
                        celda.setCellValue(texto);
                        celda.setCellStyle(style);
                    }
                    if(j == 0){
                        i = llena(fila,i,res401,hoja,style);
                    }
                    else if(j==1){
                        i = 1;
                        i = llena(fila,i,res402,hoja,style);
                    }
                    else if(j==2){
                        i = 1;
                        i = llena(fila,i,res403,hoja,style);
                    }
                    else if(j==3){
                        i = 1;
                        i = llena(fila,i,res405,hoja,style);
                    }
                    else if(j==4){
                        i = 1;
                        i = llena(fila,i,res406,hoja,style);
                    }
                    else if(j==5){
                        i = 1;
                        i = llena(fila,i,res407,hoja,style);
                    }
                        
                }
                System.out.println(i);
                
                s401.close();
                s402.close();
                s403.close();
                s405.close();
                s406.close();
                s407.close();
            }catch(Exception ex){
                ex.printStackTrace();
            }

        }catch(Exception ex){
                ex.printStackTrace();
        }


// Se salva el libro solo si se ha grabado una celda
                
        try {
            if(i>1){
                FileOutputStream elFichero = new FileOutputStream("C:\\\\reportesSucursal\\\\prueba1.xls");
                libro.write(elFichero);
                elFichero.close();
                //envioCorreo.envio("C:\\\\reportesSucursal\\\\prueba1.xls","prueba1.xls");
                
            }
        } catch (Exception e) {
            e.printStackTrace();
        }


}

}