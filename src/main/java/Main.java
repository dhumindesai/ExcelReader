
public class Main {
    public static void main(String[] args){
         String FILE_NAME;
         int SHEET_INDEX;
        try{

            switch(args.length) {
                case 0:
                {
                    System.out.println("No Arguments given.");
                    HelperMethods.printUsage();
                    break;
                }
                case 1:
                {
                    FILE_NAME = args[0];
                    ApachePOIExcelReader.printMetadata(FILE_NAME);
                    break;
                }
                case 2:{
                    FILE_NAME = args[0];
                    SHEET_INDEX = Integer.parseInt(args[1]);
                    ApachePOIExcelReader.readWorksheet(FILE_NAME, SHEET_INDEX);
                    break;
                }
                default:{
                    System.out.println("Wrong method usage.");
                    HelperMethods.printUsage();
                    break;
                }
            }

        } catch (Exception e){
                System.out.println("Something went wrong.");
               // HelperMethods.printUsage();
        }

    }
}
