
import java.awt.Color;
import java.awt.Font;
import java.awt.Graphics;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.ArrayList;
import javax.imageio.ImageIO;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import net.glxn.qrgen.QRCode;
import net.glxn.qrgen.image.ImageType;

/**
 * Este programa fue creado con el objetivo de Generar códigos QR a partir de
 * una lista de tiendas almecenadas en un archivo de excel la cual encripta esos
 * strings en MD5 y con ese nuevo String se generan imagenes con el codigo QR
 * nuevo, asimismo se encarga de escribir la traducción a MD5 al lado del string
 * original en la hoja de excel además de poner a cada imagen QR su traducción 
 * al pie de la misma.
 *
 * @author Leonardo Orozco
 */
public class GeneradorQrMd5 {

    private final String RUTA_IMGS = "C:\\Users\\Orozc\\Desktop\\GeneradorQR\\";
    private final String RUTA_EXCEL = "C:\\Users\\Orozc\\Desktop\\GeneradorQR\\tiendas.xls";
    private ArrayList<String> tiendas;
    private final int IMG_SIZE = 98; //imagen cuadrada

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        GeneradorQrMd5 qrGen = new GeneradorQrMd5();

        qrGen.tiendas = new ArrayList<String>();

        qrGen.readExcel(qrGen.RUTA_EXCEL);
        qrGen.writeExcel(qrGen.RUTA_EXCEL);
        qrGen.makeQR(qrGen.RUTA_IMGS, qrGen.IMG_SIZE);
        qrGen.addTextToImage(qrGen.RUTA_IMGS, qrGen.IMG_SIZE);
        //qrGen.imprimeTiendas(qrGen.tiendas);
    }

    /**
     * Método para testing
     * @param arr 
     */
    public void imprimeTiendas(ArrayList<String> arr) {
        for (String elemento : arr) {
            System.out.println(elemento);
        }
    }

    /**
     * Abre un excel y copia el texto de las celdas para llenar un ArrayList
     * 
     * @param xlsPath 
     * xlsPath: ubicación del archivo de excel con nombre y extención
     * 
     * Para que este método funcione:
     * 1.- Crear un archivo de excel con una lista de strings en la primer columna de la hoja
     * 2.- Guardarlo como archivo de excel 97-200 .xls no .xlsx
     * 3.- dicho archivo debe de estar en la ubicación deseada al igual que la constante RUTA_EXCEL
     * 4.- El archivo no debe ser usado durante la ejecución de este método
     */
    public void readExcel(String xlsPath) {
        try {
            Workbook libro = Workbook.getWorkbook(new File(xlsPath));
            Sheet hoja = libro.getSheet(0);
            Cell[] columna = hoja.getColumn(0);

            for (Cell celda : columna) {
                tiendas.add(celda.getContents());
            }

            libro.close();

        } catch (IOException | BiffException ex) {
            System.err.println(ex);
        }
    }

    public void writeExcel(String xlsPath) {
        try {
            Workbook libroExistente = Workbook.getWorkbook(new File(xlsPath));
            WritableWorkbook libroEditable = Workbook.createWorkbook(new File(xlsPath), libroExistente);
            libroExistente.close();
            WritableSheet hoja = libroEditable.getSheet(0);

            WritableFont fuente = new WritableFont(WritableFont.ARIAL, 12, WritableFont.NO_BOLD);
            WritableCellFormat formato = new WritableCellFormat(fuente);

            for (int i = 0; i < tiendas.size(); i++) {
                Label label = new Label(1, i, getMD5(tiendas.get(i)), formato);
                hoja.addCell(label);
            }

            libroEditable.write();
            libroEditable.close();

        } catch (IOException | BiffException | WriteException ex) {
            System.err.println(ex);
        }
    }

    public String getMD5(String input) {
        try {
            MessageDigest md = MessageDigest.getInstance("MD5");
            byte[] messageDigest = md.digest(input.getBytes());
            BigInteger number = new BigInteger(1, messageDigest);
            String hashtext = number.toString(16);

            while (hashtext.length() < 32) {
                hashtext = "0" + hashtext;
            }
            return hashtext;
        } catch (NoSuchAlgorithmException e) {
            throw new RuntimeException(e);
        }
    }

    public void makeQR(String path, int size) {
  
        for (String tienda : tiendas) {
            String filename = path+tienda+".png";
            String md5 = getMD5(tienda);
            ByteArrayOutputStream byteArrayStream = QRCode.from(md5).to(ImageType.PNG).withSize(size, size).stream();
            try {
                FileOutputStream fout;
                fout = new FileOutputStream(new File(filename));
                fout.write(byteArrayStream.toByteArray());
                fout.flush();
                fout.close();
            } catch (FileNotFoundException e) {
                System.err.println(e);
            } catch (IOException e) {
                System.err.println(e);
            }
        }

    }

    public void addTextToImage(String path, int size) {
        float fSize = (float)size;
        int x = Math.round(fSize/8);
        int y = Math.round(fSize-(fSize/16));
        
        for(String tienda : tiendas){
            String filename = path+tienda+".png";
            
            try {
            BufferedImage img = ImageIO.read(new File(filename));

            Graphics g = img.getGraphics();
            g.setFont(new Font("Arial", Font.BOLD, 11));
            g.setColor(Color.BLACK);
            g.drawString(tienda, x, y);//posicion del texto x,y
            g.dispose();

            ImageIO.write(img, "png", new File(filename));
            img.flush();
        } catch (IOException ex) {
            System.err.println(ex);
        }
        }
    }

}
