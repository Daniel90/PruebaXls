/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package my.pruebaxls;

/**
 *
 * @author Daniel Abrilot
 */

import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.swing.JOptionPane;

public class envioCorreo {
    public static void envio(String path,String nombre){
        try{
            Properties props = new Properties();
            props.put("mail.smtp.host", "smtp.tie.cl");
            
            props.setProperty("mail.smtp.starttls.enable", "true");
            props.setProperty("mail.smtp.port", "25");
            props.setProperty("mail.smtp.user", "informatica@mhochschild.cl");
            props.setProperty("mail.smtp.auth", "true");
            
            //Texto
            Session session = Session.getDefaultInstance(props, null);
            BodyPart texto = new MimeBodyPart();
            texto.setText("Pruebas xls automatico en Windows");
            
            //Adjunto
            BodyPart adjunto = new MimeBodyPart();
            adjunto.setDataHandler(
                new DataHandler(new FileDataSource(path)));
            adjunto.setFileName(nombre);
            
            // Una MultiParte para agrupar texto e imagen.
            MimeMultipart multiParte = new MimeMultipart();
            multiParte.addBodyPart(texto);
            multiParte.addBodyPart(adjunto);
            
            // Se compone el correo, dando to, from, subject y el
            // contenido.
            MimeMessage message = new MimeMessage(session);
            message.setFrom(new InternetAddress("informatica@mhochschild.cl"));
            message.setRecipients(Message.RecipientType.TO, InternetAddress.parse("informatica@mhochschild.cl"));
            
            message.setSubject("Reporte Hochschild para alertas");
            message.setContent(multiParte);

            // Se envia el correo.
            Transport t = session.getTransport("smtp");
            t.connect("informatica@mhochschild.cl", "inf01200");
            //if(stockVirtual<MinimoPermitido)
            //Enviar correo con alerta
            //else
            //Borrar porque no entra en proceso
            t.sendMessage(message, message.getAllRecipients());
            t.close();
            //JOptionPane.showMessageDialog(null,"El archivo fue enviado al correo");
            System.exit(0);
        }
        catch(Exception e){
            //JOptionPane.showMessageDialog(null, "El archivo no fue enviado al correo!","!",JOptionPane.ERROR_MESSAGE);
            e.printStackTrace();
            System.exit(0);
        }
    }
}
