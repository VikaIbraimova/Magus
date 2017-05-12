package ru.d1g;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.context.ApplicationContext;
import org.springframework.context.support.ClassPathXmlApplicationContext;
import ru.d1g.exceptions.ParseException;

import javax.swing.*;
import java.time.Duration;
import java.time.Instant;

/**
 * Created by A on 03.05.2017.
 */
public class Main {

    private static Logger log = LoggerFactory.getLogger(Main.class);
    private static ApplicationContext applicationContext = new ClassPathXmlApplicationContext("spring-config.xml");
    private static Parser parser = applicationContext.getBean(Parser.class);

    public static void main(String[] args) {
        try {
            Instant start = Instant.now();
            log.debug("parsing started");
            parser.parse();
            Instant end = Instant.now();
            log.debug("parsing time: {}", Duration.between(start,end));
        } catch (ParseException exception) {
            String message = exception.getMessage();
            if (exception.getCause() != null) {
                message = message + exception.getCause();
            }
            JOptionPane.showMessageDialog(null, "Parse exception:\n" + message);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
    }
}