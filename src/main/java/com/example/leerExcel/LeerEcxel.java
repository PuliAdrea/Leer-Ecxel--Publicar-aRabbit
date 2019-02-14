package com.example.leerExcel;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.amqp.rabbit.core.RabbitTemplate;
import org.springframework.boot.CommandLineRunner;
import org.springframework.stereotype.Component;
import org.apache.commons.io.FilenameUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;


@Component
public class LeerEcxel implements CommandLineRunner {
    public static  RabbitTemplate rabbitTemplate;
    Logger logger = LoggerFactory.getLogger(LoggingController.class);
    public LeerEcxel(RabbitTemplate rabbitTemplate) {

        this.rabbitTemplate = rabbitTemplate;

    }

    @Override
    public void run(String... args) throws Exception {
        try (
                InputStream url= AcederALaRuta();

                Workbook workbook = StreamingReader.builder()
                        .rowCacheSize(70000)
                        .bufferSize(4096)
                        .open(url)) {
            Sheet hoja = ObtenerHoja(workbook);
            Row cabezeras = LeerPrimeraFila(hoja);
            List<String> titulos = ObtenerPrimeraFilaComoLista(cabezeras);
            JsonNode nodo = CrearNodoJson();

            int i= 0;
            String json =" ";

            for (Row fila : hoja) {
                for (Cell celda : fila) {
                    ((ObjectNode) nodo).put( titulos.get(i), celda.getStringCellValue());
                    i++;
                }
                json =  nodo.toString();
                i=0;
                logger.info(json);
                publicarARabbit(json);
            }
        }catch ( Exception e){
            logger.error(e.getMessage(),e);
        }
    }



    private void publicarARabbit(String json) {
        rabbitTemplate.convertAndSend(Conexion.topicExchangeName, "foo.bar.baz", json);
    }

    public JsonNode CrearNodoJson() {
        ObjectMapper mapper = new ObjectMapper();
        JsonNode nodo = mapper.createObjectNode();
        logger.info("JsonNode Construido" );
        return nodo;
    }

    public List<String> ObtenerPrimeraFilaComoLista(Row cabezeras) {

        List<String> titulos = new ArrayList<String>();
        int i=0;
        for (Cell c : cabezeras) {
            titulos.add(cabezeras.getCell(i).getStringCellValue());
            i++;
        }
        logger.info("Se encontraron los titulos:" + titulos);
        return  titulos;
    }
    public Row LeerPrimeraFila(Sheet hoja) {
        Row cabezeras = hoja.rowIterator().next();
        logger.info("se realizo la lectura de la primera fila");
        return  cabezeras;
    }

    public Sheet ObtenerHoja(Workbook workbook) {
        Sheet hoja = workbook.getSheetAt(0);
        logger.info("se obtuvo la  "+ hoja.getSheetName());

        return hoja;
    }
    public InputStream AcederALaRuta() throws FileNotFoundException {
        String archivo ="base1.xlsx";

        InputStream url = new FileInputStream(new File
                ("C:\\Users\\ypulido\\Desktop\\"+archivo));
        logger.info("leyendo archivo:"+ archivo );
        return url;

    }

}
