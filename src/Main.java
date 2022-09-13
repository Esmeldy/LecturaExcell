import org.apache.poi.ss.usermodel.*;


import java.io.*;
import java.util.*;


public class Main {


    public static void main(String[] args) {
       try {
           //Primero La ruta del archivo para poder leerlo
           String File = "Nombres.xlsx";
           //Creamos un libro de trabajo
           Sheet sh;
           try (Workbook wb = WorkbookFactory.create(new File(File))) {
               //En este caso se usa para iterar las distintas hojas que tenga el libro
               sh = wb.getSheetAt(0);
           }

           //En este caso necesitamos la columna de Email, que es la 1, ya que la columna
           //de los nombres es la 0
           //para obtener esos datos creamos el iterador de tipo columna (Row)
           Iterator<Row> col = sh.rowIterator();
           //Creamos la colección para almacenar el resultado
           //en este caso se usa HashSet para evitar duplicated
           Set<String> emails = new HashSet<>();

           while (col.hasNext()) {
               Row row = col.next();

               //Para evitar introducir el título de la primera fila,
               //Se crea el condicional
               if (!row.getCell(1).getStringCellValue().equalsIgnoreCase("Email")){
                   emails.add(row.getCell(1).getStringCellValue());
               }

           }
           System.out.println("El archivo excel contiene los siguientes Emails: ");
           for (String email : emails) {
               System.out.println(" "+email);
           }
           System.out.println("En total: "+emails.size());
       }
       catch (IOException e) {
           System.out.println(e.getMessage());
       }
    }
}
