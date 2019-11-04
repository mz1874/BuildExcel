public class TestExcelUtil {

    private static Logger logger = LoggerFactory.getLogger(TestExcelUtil.class);

    public static void main(String[] args) throws FileNotFoundException {
        importFile();
    }

    public static void importFile() {
        File file = new File(".\\wd.xlsx");
        System.out.println(file.exists());
        InputStream inputStream = null;
        try {
            inputStream = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            logger.error("文件无法找到--->{}", e);
        }
        XSSFWorkbook workbook = null;
        try {
            workbook = new XSSFWorkbook(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
            logger.error("无法创建工作对象------>{}", e);
        }
        int sheets = workbook.getNumberOfSheets();
        boolean isHaveNeedSheet = false;
        logger.info("显示工作表的表空间------>{}", sheets);
        for (int i = 0; i < sheets; i++) {
            XSSFSheet sheetAt = workbook.getSheetAt(i);
            String sheetName = sheetAt.getSheetName();
            if (sheetName == null || !sheetName.equals("化学品询价")) {
                continue;
            }
            isHaveNeedSheet = true;
            int rows = sheetAt.getPhysicalNumberOfRows();
            logger.info("[当前工作表具有的行数] ----->{}", rows);
            int cell = sheetAt.getRow(0).getPhysicalNumberOfCells();
            logger.info("[当前工作表的总行数] ------>{}", cell);
            logger.info("开始解析excel ------------- start");
            Long StartTime = System.currentTimeMillis();
            for (int j = 0; j < rows; j++) {
                if (j == 0) {
                    logger.info("跳过首行--------------------------------------------------");
                    continue;
                }
                for (int cellRow = 0; cellRow < cell; cellRow++) {
                    CellType cellTypeEnum = sheetAt.getRow(j).getCell(cellRow).getCellTypeEnum();
                    switch (cellTypeEnum) {
                        case _NONE:
                            logger.info("***第 {}----行，第  {} -----列 ，数据--->{}", j, cellRow, "没有一个");
                            break;
                        case ERROR:
                            logger.info("***第 {}----行，第  {} -----列 ，数据--->{}", j, cellRow, "错误");
                            break;
                        case STRING:
                            String stringCellValue = sheetAt.getRow(j).getCell(cellRow).getStringCellValue();
                            logger.info("***第 {}----行，第  {} -----列 ，数据--->{}", j, cellRow, stringCellValue);
                            break;
                        case BLANK:
                            logger.info("***第 {}----行，第  {} -----列 ，数据--->{}", j, cellRow, "空白");
                            break;
                        case BOOLEAN:
                            boolean booleanCellValue = sheetAt.getRow(j).getCell(cellRow).getBooleanCellValue();
                            logger.info("***第 {}----行，第  {} -----列 ，数据--->{}", j, cellRow, booleanCellValue);
                            break;
                        case NUMERIC:
                            double numericCellValue = sheetAt.getRow(j).getCell(cellRow).getNumericCellValue();
                            logger.info("***第 {}----行，第  {} -----列 ，数据--->{}", j, cellRow, numericCellValue);
                            break;
                        default:
                            break;
                    }
                }
                System.out.println("-------------------------------------------------------------------------------------");
            }
        }
    }


}
