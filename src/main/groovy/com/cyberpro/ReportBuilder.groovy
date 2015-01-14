package com.cyberpro

import net.sf.jxls.transformer.Configuration
import net.sf.jxls.transformer.XLSTransformer
import org.apache.poi.ss.usermodel.Workbook

/**
 * Created by hlinov on 14.01.15.
 */
class ReportBuilder {
    def byte[] buildReport(){
        Map<String, Object> beans = new HashMap<String, Object>();
        beans.put("reportItem", getData());
        XLSTransformer transformer = new XLSTransformer();

        def configuration = new Configuration()
        configuration.addExcludeSheet("Sheet2")
        transformer.setConfiguration(configuration)
        transformer.markAsFixedSizeCollection("reportItem");

        InputStream inputStream = this.getClass().getResourceAsStream("/template.xlsx");
        Workbook workbook = transformer.transformXLS(inputStream, beans);
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        return outputStream.toByteArray();
    }
    def List<ReportRow> getData(){
        return [
                new ReportRow(year: 2014, sum: 100),
                new ReportRow(year: 2015, sum: 101),
                new ReportRow(year: 2015, sum: 101)
        ];
    }
    def public static void main(String[] args){
        def result = new File("test.xlsx")
        OutputStream out = new FileOutputStream(result);
        out.write(new ReportBuilder().buildReport());
        out.close()
        println("Result saved into $result.canonicalPath")
    }
}
