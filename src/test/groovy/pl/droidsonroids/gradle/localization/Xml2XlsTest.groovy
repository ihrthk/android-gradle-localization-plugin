import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.junit.Test

class Xml2XlsTest extends GroovyTestCase {

    static File TEST_RES_DIR = new File("src/test/resources/pl/droidsonroids/gradle/localization");

    @Test
    void testXml2Xls() {

//        def map = getMap(new File(getClass().getResource('res').getPath()))

//        writer(file, map)
//        def lanuages = ['', 'cs', 'de', 'es', 'fr', 'hu', 'it', 'ja', 'ko', 'nl', 'pl', 'pt-rBR',
//                        'ru', 'sv', 'zh-rCN', 'zh-rTW'] as String[]
        def lanuages = ['', 'ug'] as String[]
        def defLanuage = "en";
        def sourceDir = new File(TEST_RES_DIR, "res")
        def files = buildResDir(sourceDir, lanuages)
//      Map<String, HashMap<String, String>>
        def map = getMap(files);

        def outFile = new File(TEST_RES_DIR, "res/lite-cmxj.xls")
        writer(outFile, "sheet1", defLanuage, lanuages, map)
        println("total size:" + map.size())

    }

    private static File[] buildResDir(dir, strs) {
        File[] files = new File[strs.length]
        int i = 0;
        strs.each {
            def s = ''.equals(it) ? "values" : "values-" + it
            def var = dir.getPath() + "/" + s
            File file = new File(var, "strings.xml")
            if (file.exists()) {
                files[i] = file
            } else {
                throw new IllegalArgumentException(file.getAbsolutePath() + " not exist")
            }
            i++
        }
        return files
    }

    private static Map<String, HashMap<String, String>> getMap(File[] files) {
        Map<String, HashMap<String, String>> result = new LinkedHashMap<String, HashMap<String, String>>()
        files.each {
            String fileName = it.getParentFile().getName()
            new XmlParser().parse(it).each {
                def name = it.attributes().get('name')
                def value = it.value().text();

                //add tag
                if (fileName.equals('values') && !result.containsKey(name)) {
                    HashMap<String, String> map = new HashMap<String, String>()
                    result.put(name, map)
                }

                def map = result.get(name)
                map.put(fileName, value)
            }
        }


        result.each { key, value ->
            if (value.size() == 1) {
                value.put("values-zh-rCN", "")
            }
        }
        return result;
    }

    private
    static void writer(File file, String sheet, String defLanuage, String[] lanuages, Map<String, HashMap<String, String>> map) throws IOException {
        //1.create workbook object
        Workbook workbook = file.getAbsolutePath().endsWith("xls") ?
                new HSSFWorkbook() : new XSSFWorkbook();
        //2.create sheet object
        Sheet sheet1 = (Sheet) workbook.createSheet(sheet);
        //3.for each row,write data
        Row r = (Row) sheet1.createRow(0);

        r.createCell(0).setCellValue("Android")
        for (int i = 0; i < lanuages.length; i++) {
            String language = "".equals(lanuages[i]) ? defLanuage : lanuages[i];
            def folderName = "values" + ("".equals(lanuages[i]) ? "" : "-" + lanuages[i])
            String percent = getLength(map, folderName) + "/" + map.size();
            r.createCell(i + 1).setCellValue(language + "(" + percent + ")")
        }

        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.RED.getIndex());
        style.setFillBackgroundColor(IndexedColors.RED.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);

        CellStyle styleAlignLeft = workbook.createCellStyle();
        styleAlignLeft.setAlignment(CellStyle.ALIGN_LEFT)

        int i = 1
        map.each {
            Row row = (Row) sheet1.createRow(i);
            row.createCell(0).setCellValue(it.key)
            int j = 1
            it.value.each {
                def cell = row.createCell(j)
                cell.setCellValue(it.value)
                cell.setCellStyle(styleAlignLeft);
                if ("".equals(it.value)) {
                    cell.setCellStyle(style);
                }
                j++
            }
            i++
        }

        //4.create file stream
        OutputStream stream = new FileOutputStream(file);
        //5.write data
        workbook.write(stream);
        //6.close stream
        stream.close();
    }

    public static int getLength(Map<String, HashMap<String, String>> map, String dirName) {
        int count = 0;
        map.each { key, v ->
            if (!"".equals(v.get(dirName))) {
                count++;
            }
        }
        return count;
    }
}