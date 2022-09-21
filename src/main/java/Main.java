public class Main {
    public static void main(String[] args) {

        String fileName = "domaci22.xlsx";

        ApachePoiUtils.readAndPrint(fileName);

        ApachePoiUtils.write(fileName);

        ApachePoiUtils.readAllAndPrint(fileName);



    }
}
