
package codigo;



import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import static javax.swing.WindowConstants.EXIT_ON_CLOSE;


public class PDF2 extends JFrame {
    
    private JPanel panel;
    private JButton boton;
    private JLabel texto;

    public PDF2() {

        this.setSize(450, 250);
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setTitle("E X P O R T A R");
        setLocationRelativeTo(null);
        setMinimumSize(new Dimension(200, 200));

        Iniciar();

    }

    private void Iniciar() {

        Panel();
        Pregunta();
        Boton();

    }

    private void Panel() {

        panel = new JPanel();
        panel.setLayout(null);
        this.getContentPane().add(panel);
    }

     public void Pregunta() {

        texto = new JLabel();
        texto.setBounds(130, 20, 400, 50);
        texto.setText("Se ha creado el archivo PDF.");
        panel.add(texto);

    }
     
    private void Boton() {

        boton = new JButton("Abrir");
        boton.setBounds(180, 120, 90, 40);
        boton.setForeground(java.awt.Color.blue);
        boton.setFont(new java.awt.Font("Baskerville", java.awt.Font.CENTER_BASELINE, 15));
        boton.setEnabled(true);

        panel.add(boton);
        
        ActionListener accion;
        accion = new ActionListener() {
            @Override

            public void actionPerformed(ActionEvent e) {

                try {
                    File path = new File("/Users/1/Desktop/ejemplo.pdf");
                    Desktop.getDesktop().open(path);
                } catch (IOException ex) {
                    ex.printStackTrace();
                }

            }
        };

        boton.addActionListener(accion);

    }

    
    
}
