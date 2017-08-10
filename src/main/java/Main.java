
public class Main {
    public static void main(String[] args){
        // Commit change
        try{
            final String FILE_NAME = args[0];
            final int SHEET_INDEX = Integer.parseInt(args[1]);
            ApachePOIExcelReader.readExcel(FILE_NAME, SHEET_INDEX);
        } catch (IndexOutOfBoundsException e){
            if (args.length == 0){
                System.out.println("No arguments are given.\n");
                HelperMethods.printUsage();
            } else{
                System.out.println("No worksheet index provided. Please provide valid worksheet index.");
                HelperMethods.printUsage();
            }
        }

    }
}
