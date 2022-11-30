package codigo;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.swing.JButton;
import javax.swing.JFrame;
import static javax.swing.JFrame.EXIT_ON_CLOSE;
import javax.swing.JLabel;
import javax.swing.JPanel;
import static javax.swing.WindowConstants.EXIT_ON_CLOSE;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Generar extends JFrame {

    private JPanel panel;

    public Generar() {

        this.setSize(550, 350);
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setTitle("E X P O R T A R");
        setLocationRelativeTo(null);
        setMinimumSize(new Dimension(200, 200));

        Iniciar();

    }

    private void Iniciar() {

        Panel();
        Etiquetas();
        BotonExcel();
        BotonPDF();
    }

    private void Panel() {

        panel = new JPanel();
        panel.setLayout(null);
        this.getContentPane().add(panel);
    }

    private void Etiquetas() {

        JLabel etiqueta = new JLabel("¿En donde deseas exportar?");
        panel.add(etiqueta);
        etiqueta.setBounds(100, 70, 500, 80);
        etiqueta.setForeground(Color.black);
        etiqueta.setFont(new Font("Baskerville", Font.ITALIC, 25));

    }

    private void BotonExcel() {
        JButton boton = new JButton("Excel");
        boton.setBounds(90, 250, 150, 40);
        boton.setForeground(Color.DARK_GRAY);
        boton.setFont(new Font("Baskerville", Font.ITALIC, 15));
        boton.setEnabled(true);
        panel.add(boton);

        boton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent arg0) {

                Workbook libro = new XSSFWorkbook();
                final String nombreArchivo = ".xlsx";
                Sheet hoja = libro.createSheet("Hoja 1");
                Row primeraFila = hoja.createRow(0);
                Cell primeraCelda = primeraFila.createCell(0);
                primeraCelda.setCellValue("UbicaciÃ³n-la primera celda y primera fila");
                File directorioActual = new File("/Users/1/Desktop/ejemplo");
                String ubicacion = directorioActual.getAbsolutePath();
                String ubicacionArchivoSalida = ubicacion.substring(0, ubicacion.length() - 1) + nombreArchivo;
                FileOutputStream outputStream;
                try {
                    outputStream = new FileOutputStream(ubicacionArchivoSalida);
                    libro.write(outputStream);
                    libro.close();
                    System.out.println("Libro guardado correctamente");
                } catch (FileNotFoundException ex) {
                    System.out.println("Error de filenotfound");
                } catch (IOException ex) {
                    System.out.println("Error de IOException");
                }

                EXCEL excel = new EXCEL();
                excel.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
                excel.setVisible(true);

                setVisible(false);

            }
        });

    }

    private void BotonPDF() {
        JButton boton = new JButton("PDF");
        boton.setBounds(300, 250, 150, 40);
        boton.setForeground(Color.darkGray);
        boton.setFont(new Font("Baskerville", Font.ITALIC, 15));
        boton.setEnabled(true);
        panel.add(boton);

        boton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent arg0) {

 
                String imagen = "/Users/1/Desktop/imagen.jpeg";
                try {
                    PDDocument documento = new PDDocument();
                    PDPage pagina = new PDPage(PDRectangle.LETTER);
                    documento.addPage(pagina);

                    PDImageXObject image = PDImageXObject.createFromFile(imagen, documento);
                    PDPageContentStream contenido = new PDPageContentStream(documento, pagina);

                    //contenido.drawImage(image, 20f, 20f);

                    contenido.drawImage(image, 40, 40, 550, 550);
                    
                    contenido.beginText();
                    contenido.setFont(PDType1Font.TIMES_ROMAN, 17);
                    contenido.setLeading(14.5f);
                    contenido.newLineAtOffset(20, pagina.getMediaBox().getHeight() - 52);
                    contenido.showText("PRUEBA DE EXPORTACIOON PDF");
                    
                    contenido.endText();

                    contenido.close();

                    documento.save("/Users/1/Desktop/ejemplo.pdf");

                } catch (Exception x) {
                    System.out.println("Error: " + x.getMessage().toString());
                }

                /////////////////////////////////////////////////////////////////   
                PDF2 p = new PDF2();
                p.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
                p.setVisible(true);

                setVisible(false);

            }
        });

    }
}



