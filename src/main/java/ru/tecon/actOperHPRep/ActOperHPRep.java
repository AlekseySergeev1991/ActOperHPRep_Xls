package ru.tecon.actOperHPRep;

import org.apache.commons.codec.DecoderException;
import org.apache.commons.codec.binary.Hex;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import ru.tecon.actOperHPRep.model.RepType;
import ru.tecon.actOperHPRep.model.ReportObject;
import ru.tecon.actOperHPRep.model.Value;

import javax.sql.DataSource;
import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.sql.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

public class ActOperHPRep {
    private static final Logger LOGGER = Logger.getLogger(ActOperHPRep.class.getName());
    private static final String LOAD_REP_TYPE = "select a.*,b.par_code from admin.REP_FACT_REPORT a, admin.dz_param b where a.par_id=b.id and a.id=?";
    private static final String LOAD_OBJECT = "select * from dsp_0079t.sel_object_list(?)";
    private static final String LOAD_VALUE_T1_TYPE = "select * from dsp_0079t.sel_rep_fact_t(?,?)";
    private static final String LOAD_VALUE_T7_TYPE = "SELECT * from dsp_0079t.sel_rep_fact_t7(?,?)";
    private static final String LOAD_VALUE_V_TYPE = "SELECT * from dsp_0079t.sel_rep_fact_v(?, ?) order by zone, measure, time_stamp";
    private static final String LOAD_VALUE_P_TYPE = "SELECT * from dsp_0079t.sel_rep_fact_p(?, ?)";
    private static final String LOAD_VALUE_G_TYPE = "SELECT * from dsp_0079t.sel_rep_fact_g(?, ?)";

    private static final String INTERRUPTED = "select  dsp_0079t.get_rep_status(?)";
    private static final String PERCENT = "call dsp_0079t.update_percent(?, ?)";
    private static final String DELSQL = "call dsp_0079t.del_report(?,?)";
    private static final String SQL = "select * from dsp_0079t.save_report(?,?)";
    private static final String FINSQL = "call dsp_0079t.update_status(?, 'F')";



    private int Rows;
    private HashMap<String, CellStyle> colors = new HashMap<>();
    private DataSource dsR;
    public void setDsR(DataSource dsR) {
        this.dsR = dsR;
    }

    private DataSource dsRW;

    public void setDsRW(DataSource dsRW) {
        this.dsRW = dsRW;
    }

    public static void makeReport (int repId, DataSource dsR, DataSource dsRW) throws InterruptedException {
        long currentTime = System.nanoTime();
        LOGGER.log(Level.INFO, "start make report {0}", repId);

        ActOperHPRep ar = new ActOperHPRep();
        ar.setDsR(dsR);
        ar.setDsRW(dsRW);
        SXSSFWorkbook w;
        try {
            w = ar.printReport(repId, dsR, dsRW);
            ar.saveReportIntoTable (w, repId, dsRW);
//            ar.saveReportIntoFile(w, "C:\\abc\\ActOperMOEK_DelTest.xlsx");

        } catch (IOException | SQLException | ParseException | DecoderException e) {
            LOGGER.log(Level.WARNING, "makeReport error", e);
//            e.printStackTrace();
        }
//        int j = saveReportIntoFile(w,"D:\\TEST.xls");

        LOGGER.log(Level.INFO, "report created {0} created time {1}", new java.lang.Object[]{repId, (System.nanoTime() - currentTime)});

    }

    private int saveReportIntoFile (SXSSFWorkbook workbook, String file) {

        FileOutputStream fos;
        int res = 0;

        try {
            fos = new FileOutputStream(file);
            workbook.write(fos);
            System.out.println("saveReportIntoFile");

        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Cant create file");
            System.out.println(e.getMessage());
            res = 1;
        }

        System.out.println("ok report");
        return res;
    }

    /*
      Метод создает нужный воркбук. Параметры:
      Rep_Id - идентификатор отчета
      Rep_Type - тип репорта. Принимает на вход один символ, малая латинская буква: h - часовой, d - дневной, m - месячный
      Beg_Date - Начальная дата в ткстовом формате (тут пишу в ораклиной нотации) dd-mm-yyyy hh24:mi
      End_Date - Конечная дата в ткстовом формате (тут пишу в ораклиной нотации) dd-mm-yyyy hh24:mi. Прекрасно понимаю, что ее можно рассчитать
                   с помощью количества колонок, типа отчета и начальной даты
      Rows - количество строк в отчете
      Data_Cols - количество колонок - значений параметров. В шапке на 4 колонки больше

    */
    public SXSSFWorkbook printReport (int repId, DataSource dsR, DataSource dsRW) throws IOException, SQLException, ParseException, DecoderException {
        SXSSFWorkbook wb = new SXSSFWorkbook();
        SXSSFSheet sh = wb.createSheet("Отчет");
        CellStyle headerStyle = setHeaderStyle(wb);
        CellStyle headerStyleNoBold = setHeaderStyleNoBold(wb);
        CellStyle nowStyle = setCellNow (wb);
        CellStyle tableHeaderStyle = setTableHeaderStyle(wb);


        //запрос для получения информации о типе отчета диапоазоне дат
        RepType repType = loadRepType(repId, dsRW);

        if (repType.getTypeCode().equals("Gт")) {

            SXSSFRow row_1 = sh.createRow(0);
            row_1.setHeight((short) 435);
            SXSSFCell cell_1_1 = row_1.createCell(0);
            cell_1_1.setCellValue("ПАО \"МОЭК\": АС \"ТЕКОН - Диспетчеризация\"");

            CellRangeAddress title = new CellRangeAddress(0, 0, 0, 4);
            sh.addMergedRegion(title);
            cell_1_1.setCellStyle(headerStyle);

            SXSSFRow row_2 = sh.createRow(1);
            row_2.setHeight((short) 435);
            SXSSFCell cell_2_1 = row_2.createCell(0);
            cell_2_1.setCellValue("Анализ фактической работы ЦТП по показателю: G1 - G2" );
            CellRangeAddress formName = new CellRangeAddress(1, 1, 0, 4);
            sh.addMergedRegion(formName);
            cell_2_1.setCellStyle(headerStyle);

            SXSSFRow row_3 = sh.createRow(2);
            row_3.setHeight((short) 435);
            SXSSFCell cell_3_1 = row_3.createCell(0);
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd.MM.yyyy");
            LocalDateTime begFormatted = repType.getBeg();
            String stringBeg = begFormatted.format(formatter);
            LocalDateTime endFormattedT1 = repType.getEnd();
            String stringEnd = endFormattedT1.format(formatter);
            cell_3_1.setCellValue("за период: " + stringBeg + " - " + stringEnd);
            cell_3_1.setCellStyle(headerStyleNoBold);
            CellRangeAddress period = new CellRangeAddress(2, 2, 0, 4);
            sh.addMergedRegion(period);
            cell_3_1.setCellStyle(headerStyleNoBold);

            SXSSFRow row_4 = sh.createRow(3);
            row_4.setHeight((short) 435);
            SXSSFCell cell_4_1 = row_4.createCell(0);
            String interval = "";
            switch (repType.getInterval()) {
                case ("D"):
                    interval = "Часовой";
                    break;
                case ("M"):
                    interval = "Дневной";
                    break;
            }
            cell_4_1.setCellValue("Интервал: " + interval);
            CellRangeAddress intervalMerge = new CellRangeAddress(3, 3, 0, 4);
            sh.addMergedRegion(intervalMerge);
            cell_4_1.setCellStyle(headerStyle);

            // Печатаем отчетов зад общий для всех отчетов
            String now = new SimpleDateFormat("dd.MM.yyyy HH:mm").format(new Date());
            SXSSFRow row_5 = sh.createRow(4);
            row_5.setHeight((short) 435);
            SXSSFCell cell5_1 = row_5.createCell(0);
            cell5_1.setCellStyle(nowStyle);
            cell5_1.setCellValue("Отчет сформирован  " + now);
            CellRangeAddress nowDone = new CellRangeAddress(4, 4, 0, 4);
            sh.addMergedRegion(nowDone);
        } else {
            setHeader(sh, headerStyle, headerStyleNoBold, nowStyle, repType);
        }

        // Пожалуй, наполню-ка я лист отдельным методом. И сначала заполняем его, чтобы узнать Rows. Cols
        switch (repType.getTypeCode()) {
            case ("Tт"):
            case ("Tто"):
            case ("Tц"):
            case ("Tцо"):

                int begRowT1 = 9;  // строка в екселе, с которой начинается собственно отчет.
                List<LocalDateTime> dateListT1 = new ArrayList<>();
                LocalDateTime localDateTempT1 = repType.getBeg();

                dateListT1.add(localDateTempT1);
                //-- Считаем количество дней между датами

                int colsT1 = fillDateList(repType, dateListT1, localDateTempT1);

                //Приступим к основному отчету
                // Устанавливаем ширины колонок. В конце мероприятия
                setColumnWidth(sh);

                for (int i = 1; i <= colsT1; i++) {
                    sh.setColumnWidth(5*i+1, 4669);
                    sh.setColumnWidth(5*i+2, 4669);
                    sh.setColumnWidth(5*i+3, 4695);
                    sh.setColumnWidth(5*i+4, 4695);
                    sh.setColumnWidth(5*i+5, 4659);
                }

                // Прекрасно, Заголовок сделали. Готовим шапку.
                SXSSFRow row_6T1 = sh.createRow(5);
                row_6T1.setHeight((short) 350);
                SXSSFRow row_7T1 = sh.createRow(6);
                row_7T1.setHeight((short) 2370);
                SXSSFRow row_8T1 = sh.createRow(7);
                row_8T1.setHeight((short) 350);
                SXSSFRow row_9T1 = sh.createRow(8);
                row_9T1.setHeight((short) 350);


                SXSSFCell cell_6_1T1 = row_6T1.createCell(0);
                cell_6_1T1.setCellStyle(tableHeaderStyle);
                cell_6_1T1.setCellValue("№ п/п");


                SXSSFCell cell_6_2T1 = row_6T1.createCell(1);
                cell_6_2T1.setCellStyle(tableHeaderStyle);
                cell_6_2T1.setCellValue("Объект");


                SXSSFCell cell_6_3T1 = row_6T1.createCell(2);
                cell_6_3T1.setCellStyle(tableHeaderStyle);
                cell_6_3T1.setCellValue("Филиал");


                SXSSFCell cell_6_4T1 = row_6T1.createCell(3);
                cell_6_4T1.setCellStyle(tableHeaderStyle);
                cell_6_4T1.setCellValue("Предприятие");


                SXSSFCell cell_6_5T1 = row_6T1.createCell(4);
                cell_6_5T1.setCellStyle(tableHeaderStyle);
                cell_6_5T1.setCellValue("Адрес ЦТП");


                SXSSFCell cell_6_6T1 = row_6T1.createCell(5);
                cell_6_6T1.setCellStyle(tableHeaderStyle);
                cell_6_6T1.setCellValue("Зона");

                createMergeForTableHeader(sh);

                // декларируем переменные для шапки

                //тут будет цикл для шапок в зависимости от количества timestamp
                int iT1 = 5; //вообще последний неизменный столбец - 6, но считаем с 0, потому 5

                for (LocalDateTime localDateTime: dateListT1) {

                    DateTimeFormatter dtfT1 = DateTimeFormatter.ofPattern("dd.MM.yyyy");
                    String timestampT1 = localDateTime.format(dtfT1);

                    SXSSFCell cell_6_7T1 = row_6T1.createCell(iT1+1);
                    cell_6_7T1.setCellStyle(tableHeaderStyle);
                    cell_6_7T1.setCellValue(timestampT1);
                    CellRangeAddress headerDateT1 = new CellRangeAddress(5, 5, iT1+1, iT1+5);
                    sh.addMergedRegion(headerDateT1);
                    CellRangeAddress borderForDateT1 = new CellRangeAddress(5, 5, iT1+1, iT1+5);
                    setBorders(borderForDateT1, sh);


                    SXSSFCell cell_7_7T1 = row_7T1.createCell(iT1+1);
                    cell_7_7T1.setCellStyle(tableHeaderStyle);
                    cell_7_7T1.setCellValue("Среднесуточная температура наружного воздуха по данным датчика ЦТП");
                    CellRangeAddress headerSensorTempT1 = new CellRangeAddress(6, 7, iT1+1, iT1+1);
                    sh.addMergedRegion(headerSensorTempT1);
                    CellRangeAddress borderForSensorTempT1 = new CellRangeAddress(6, 7, iT1+1, iT1+1);
                    setBorders(borderForSensorTempT1, sh);


                    SXSSFCell cell_7_8T1 = row_7T1.createCell(iT1+2);
                    cell_7_8T1.setCellStyle(tableHeaderStyle);
                    cell_7_8T1.setCellValue("Среднесуточная температура наружного воздуха по данным ГМЦ");
                    CellRangeAddress headerGMTempT1 = new CellRangeAddress(6, 7, iT1+2, iT1+2);
                    sh.addMergedRegion(headerGMTempT1);
                    CellRangeAddress borderForGMTempT1 = new CellRangeAddress(6, 7, iT1+2, iT1+2);
                    setBorders(borderForGMTempT1, sh);


                    SXSSFCell cell_7_9T1 = row_7T1.createCell(iT1+3);
                    cell_7_9T1.setCellStyle(tableHeaderStyle);
                    cell_7_9T1.setCellValue("Нормативный диапазон (нормативная уставка)");
                    CellRangeAddress headerNormRangeT1 = new CellRangeAddress(6, 6, iT1+3, iT1+4);
                    sh.addMergedRegion(headerNormRangeT1);
                    CellRangeAddress borderForNormRangeT1 = new CellRangeAddress(6, 6, iT1+3, iT1+4);
                    setBorders(borderForNormRangeT1, sh);


                    SXSSFCell cell_7_11T1 = row_7T1.createCell(iT1+5);
                    cell_7_11T1.setCellStyle(tableHeaderStyle);
                    if (repType.getTypeCode().equals("Tт")) {
                        cell_7_11T1.setCellValue("Среднесуточная температура сетевой воды в подающем трубопроводе Теплосети");
                    }
                    if (repType.getTypeCode().equals("Tто")) {
                        cell_7_11T1.setCellValue("Среднесуточная температура сетевой воды в обратном трубопроводе Теплосети");
                    }
                    if (repType.getTypeCode().equals("Tц")) {
                        cell_7_11T1.setCellValue("Среднесуточная температура  в подающем трубопроводе ЦО");
                    }
                    if (repType.getTypeCode().equals("Tцо")) {
                        cell_7_11T1.setCellValue("Среднесуточная температура  в обратном трубопроводе ЦО");
                    }

                    SXSSFCell cell_8_9T1 = row_8T1.createCell(iT1+3);
                    cell_8_9T1.setCellStyle(tableHeaderStyle);
                    cell_8_9T1.setCellValue("MIN");

                    SXSSFCell cell_8_10T1 = row_8T1.createCell(iT1+4);
                    cell_8_10T1.setCellStyle(tableHeaderStyle);
                    cell_8_10T1.setCellValue("MAX");

                    SXSSFCell cell_8_11T1 = row_8T1.createCell(iT1+5);
                    cell_8_11T1.setCellStyle(tableHeaderStyle);
                    cell_8_11T1.setCellValue(repType.getType());

                    SXSSFCell cell_9_7T1 = row_9T1.createCell(iT1+1);
                    cell_9_7T1.setCellStyle(tableHeaderStyle);
                    cell_9_7T1.setCellValue("\u2103");
                    SXSSFCell cell_9_8T1 = row_9T1.createCell(iT1+2);
                    cell_9_8T1.setCellStyle(tableHeaderStyle);
                    cell_9_8T1.setCellValue("\u2103");
                    SXSSFCell cell_9_9T1 = row_9T1.createCell(iT1+3);
                    cell_9_9T1.setCellStyle(tableHeaderStyle);
                    cell_9_9T1.setCellValue("\u2103");
                    SXSSFCell cell_9_10T1 = row_9T1.createCell(iT1+4);
                    cell_9_10T1.setCellStyle(tableHeaderStyle);
                    cell_9_10T1.setCellValue("\u2103");
                    SXSSFCell cell_9_11T1 = row_9T1.createCell(iT1+5);
                    cell_9_11T1.setCellStyle(tableHeaderStyle);
                    cell_9_11T1.setCellValue("\u2103");

                    iT1 = iT1+5;
                }

                LOGGER.log(Level.INFO, "Report head created {0}", repId);
                // Отлично. Заголовок и шапка сделаны. Идем по таблице, создаем и заполняем ячейки.

                fillSheetT1(wb, repId, begRowT1, dateListT1.size(), dsR, dsRW, repType);

                sh.createFreezePane(6, 9);

                LOGGER.log(Level.INFO, "Report body created {0}", repId);

                break;

            case ("Tг"):
            case ("Tго"):

                int begRowT7 = 9;  // строка в екселе, с которой начинается собственно отчет.
                List<LocalDateTime> dateListT7 = new ArrayList<>();
                LocalDateTime localDateTempT7 = repType.getBeg();

                dateListT7.add(localDateTempT7);
                //-- Считаем количество дней между датами

                int colsT7 = fillDateList(repType, dateListT7, localDateTempT7);

                //Приступим к основному отчету
                // Устанавливаем ширины колонок. В конце мероприятия
                setColumnWidth(sh);
                sh.setColumnWidth(6, 4695);
                sh.setColumnWidth(7, 4695);
                sh.setColumnWidth(8, 4695);
                sh.setColumnWidth(9, 4695);

                for (int i = 1; i <= colsT7; i++) {
                    sh.setColumnWidth(9+i, 4669);
                }

                // Прекрасно, Заголовок сделали. Готовим шапку.
                SXSSFRow row_6T7 = sh.createRow(5);
                row_6T7.setHeight((short) 350);
                SXSSFRow row_7T7 = sh.createRow(6);
                row_7T7.setHeight((short) 2370);
                SXSSFRow row_8T7 = sh.createRow(7);
                row_8T7.setHeight((short) 350);
                SXSSFRow row_9T7 = sh.createRow(8);
                row_9T7.setHeight((short) 350);

                SXSSFCell cell_6_1T7 = row_6T7.createCell(0);
                cell_6_1T7.setCellStyle(tableHeaderStyle);
                cell_6_1T7.setCellValue("№ п/п");

                SXSSFCell cell_6_2T7 = row_6T7.createCell(1);
                cell_6_2T7.setCellStyle(tableHeaderStyle);
                cell_6_2T7.setCellValue("Объект");

                SXSSFCell cell_6_3T7 = row_6T7.createCell(2);
                cell_6_3T7.setCellStyle(tableHeaderStyle);
                cell_6_3T7.setCellValue("Филиал");

                SXSSFCell cell_6_4T7 = row_6T7.createCell(3);
                cell_6_4T7.setCellStyle(tableHeaderStyle);
                cell_6_4T7.setCellValue("Предприятие");

                SXSSFCell cell_6_5T7 = row_6T7.createCell(4);
                cell_6_5T7.setCellStyle(tableHeaderStyle);
                cell_6_5T7.setCellValue("Адрес ЦТП");

                SXSSFCell cell_6_6T7 = row_6T7.createCell(5);
                cell_6_6T7.setCellStyle(tableHeaderStyle);
                cell_6_6T7.setCellValue("Зона");

                createMergeForTableHeader(sh);

                SXSSFCell cell_6_7T7 = row_6T7.createCell(6);
                cell_6_7T7.setCellStyle(tableHeaderStyle);
                cell_6_7T7.setCellValue("Нормативный диапазон (нормативная уставка)");
                CellRangeAddress headerNormRangeT7 = new CellRangeAddress(5, 6, 6, 7);
                sh.addMergedRegion(headerNormRangeT7);
                CellRangeAddress borderForNormRangeT7 = new CellRangeAddress(5, 6, 6, 7);
                setBorders(borderForNormRangeT7, sh);

                SXSSFCell cell_8_7T7 = row_8T7.createCell(6);
                cell_8_7T7.setCellStyle(tableHeaderStyle);
                cell_8_7T7.setCellValue("MIN");

                SXSSFCell cell_8_8T7 = row_8T7.createCell(7);
                cell_8_8T7.setCellStyle(tableHeaderStyle);
                cell_8_8T7.setCellValue("MAX");

                SXSSFCell cell_9_7T7 = row_9T7.createCell(6);
                cell_9_7T7.setCellStyle(tableHeaderStyle);
                cell_9_7T7.setCellValue("\u2103");
                SXSSFCell cell_9_8T7 = row_9T7.createCell(7);
                cell_9_8T7.setCellStyle(tableHeaderStyle);
                cell_9_8T7.setCellValue("\u2103");

                SXSSFCell cell_6_9T7 = row_6T7.createCell(8);
                cell_6_9T7.setCellStyle(tableHeaderStyle);
                cell_6_9T7.setCellValue("Допустимый диапазон (критическая уставка)");
                CellRangeAddress headerAvarRangeT7 = new CellRangeAddress(5, 6, 8, 9);
                sh.addMergedRegion(headerAvarRangeT7);
                CellRangeAddress borderForAvarRangeT7 = new CellRangeAddress(5, 6, 8, 9);
                setBorders(borderForAvarRangeT7, sh);

                SXSSFCell cell_8_9T7 = row_8T7.createCell(8);
                cell_8_9T7.setCellStyle(tableHeaderStyle);
                cell_8_9T7.setCellValue("MIN");

                SXSSFCell cell_8_10T7 = row_8T7.createCell(9);
                cell_8_10T7.setCellStyle(tableHeaderStyle);
                cell_8_10T7.setCellValue("MAX");

                SXSSFCell cell_9_9T7 = row_9T7.createCell(8);
                cell_9_9T7.setCellStyle(tableHeaderStyle);
                cell_9_9T7.setCellValue("\u2103");
                SXSSFCell cell_9_10T7 = row_9T7.createCell(9);
                cell_9_10T7.setCellStyle(tableHeaderStyle);
                cell_9_10T7.setCellValue("\u2103");


                // декларируем переменные для шапки

                int iT7 = 10; //вообще последний неизменный столбец - 6, но считаем с 0, потому 5

                for (LocalDateTime localDateTime: dateListT7) {

                    DateTimeFormatter dtf = DateTimeFormatter.ofPattern("dd.MM.yyyy");
                    String timestamp = localDateTime.format(dtf);

                    SXSSFCell cell_6_11T7 = row_6T7.createCell(iT7);
                    cell_6_11T7.setCellStyle(tableHeaderStyle);
                    cell_6_11T7.setCellValue(timestamp);

                    SXSSFCell cell_7_11T7 = row_7T7.createCell(iT7);
                    cell_7_11T7.setCellStyle(tableHeaderStyle);
                    if (repType.getTypeCode().equals("Tг")) {
                        cell_7_11T7.setCellValue("Среднесуточная температура горячей воды в подающем трубопроводе ГВС");
                    }
                    if (repType.getTypeCode().equals("Tго")) {
                        cell_7_11T7.setCellValue("Среднесуточная температура горячей воды в обратном трубопроводе ГВС");
                    }

                    SXSSFCell cell_8_11T7 = row_8T7.createCell(iT7);
                    cell_8_11T7.setCellStyle(tableHeaderStyle);
                    cell_8_11T7.setCellValue(repType.getType());

                    SXSSFCell cell_9_11T7 = row_9T7.createCell(iT7);
                    cell_9_11T7.setCellStyle(tableHeaderStyle);
                    cell_9_11T7.setCellValue("\u2103");

                    iT7++;
                }

                LOGGER.log(Level.INFO, "Report head created {0}", repId);
                // Отлично. Заголовок и шапка сделаны. Идем по таблице, создаем и заполняем ячейки.

                fillSheetT7(wb, repId, begRowT7, dateListT7.size(), dsR, dsRW, repType);

                sh.createFreezePane(10, 9);


                LOGGER.log(Level.INFO, "Report body created {0}", repId);

                break;

            case ("Qгп"):

                int begRowV = 9;  // строка в екселе, с которой начинается собственно отчет.
                List<LocalDateTime> dateListV = new ArrayList<>();
                LocalDateTime localDateTempV = repType.getBeg();

                dateListV.add(localDateTempV);
                //-- Считаем количество дней между датами

                int colsV = fillDateList(repType, dateListV, localDateTempV);

                //Приступим к основному отчету
                // Устанавливаем ширины колонок. В конце мероприятия
                setColumnWidth(sh);
                sh.setColumnWidth(6, 4295);
                sh.setColumnWidth(7, 4695);
                sh.setColumnWidth(8, 4695);


                for (int i = 1; i <= colsV; i++) {
                    sh.setColumnWidth(8+i, 4669);
                }

                // Прекрасно, Заголовок сделали. Готовим шапку.
                SXSSFRow row_6V = sh.createRow(5);
                row_6V.setHeight((short) 350);
                SXSSFRow row_7V = sh.createRow(6);
                row_7V.setHeight((short) 2370);
                SXSSFRow row_8V = sh.createRow(7);
                row_8V.setHeight((short) 350);
                SXSSFRow row_9V = sh.createRow(8);
                row_9V.setHeight((short) 350);


                SXSSFCell cell_6_1V = row_6V.createCell(0);
                cell_6_1V.setCellStyle(tableHeaderStyle);
                cell_6_1V.setCellValue("№ п/п");

                SXSSFCell cell_6_2V = row_6V.createCell(1);
                cell_6_2V.setCellStyle(tableHeaderStyle);
                cell_6_2V.setCellValue("Объект");

                SXSSFCell cell_6_3V = row_6V.createCell(2);
                cell_6_3V.setCellStyle(tableHeaderStyle);
                cell_6_3V.setCellValue("Филиал");

                SXSSFCell cell_6_4V = row_6V.createCell(3);
                cell_6_4V.setCellStyle(tableHeaderStyle);
                cell_6_4V.setCellValue("Предприятие");

                SXSSFCell cell_6_5V = row_6V.createCell(4);
                cell_6_5V.setCellStyle(tableHeaderStyle);
                cell_6_5V.setCellValue("Адрес ЦТП");

                SXSSFCell cell_6_6V = row_6V.createCell(5);
                cell_6_6V.setCellStyle(tableHeaderStyle);
                cell_6_6V.setCellValue("Зона");

                createMergeForTableHeader(sh);

                SXSSFCell cell_6_7V = row_6V.createCell(6);
                cell_6_7V.setCellStyle(tableHeaderStyle);
                cell_6_7V.setCellValue("Ед. изм.");
                CellRangeAddress headerMeasure = new CellRangeAddress(5, 8, 6, 6);
                sh.addMergedRegion(headerMeasure);
                CellRangeAddress borderForMeasure = new CellRangeAddress(5, 8, 6, 6);
                RegionUtil.setBorderBottom(BorderStyle.THICK, borderForMeasure, sh);
                RegionUtil.setBorderTop(BorderStyle.THICK, borderForMeasure, sh);
                RegionUtil.setBorderLeft(BorderStyle.THICK, borderForMeasure, sh);
                RegionUtil.setBorderRight(BorderStyle.THICK, borderForMeasure, sh);

                SXSSFCell cell_6_8V = row_6V.createCell(7);
                cell_6_8V.setCellStyle(tableHeaderStyle);
                cell_6_8V.setCellValue("Нормативный диапазон (нормативная уставка)");
                CellRangeAddress headerNormRangeV = new CellRangeAddress(5, 6, 7, 8);
                sh.addMergedRegion(headerNormRangeV);
                CellRangeAddress borderForNormRangeV = new CellRangeAddress(5, 6, 7, 8);
                setBorders(borderForNormRangeV, sh);

                SXSSFCell cell_8_8V = row_8V.createCell(7);
                cell_8_8V.setCellStyle(tableHeaderStyle);
                cell_8_8V.setCellValue("MIN");

                SXSSFCell cell_8_9V = row_8V.createCell(8);
                cell_8_9V.setCellStyle(tableHeaderStyle);
                cell_8_9V.setCellValue("MAX");

                SXSSFCell cell_9_8V = row_9V.createCell(7);
                cell_9_8V.setCellStyle(tableHeaderStyle);
                cell_9_8V.setCellValue("М.куб.");
                SXSSFCell cell_9_9V = row_9V.createCell(8);
                cell_9_9V.setCellStyle(tableHeaderStyle);
                cell_9_9V.setCellValue("М.куб.");

                // декларируем переменные для шапки

                int iV = 9; //вообще последний неизменный столбец - 6, но считаем с 0, потому 5

                for (LocalDateTime localDateTime: dateListV) {

                    DateTimeFormatter dtf = DateTimeFormatter.ofPattern("dd.MM.yyyy");
                    String timestamp = localDateTime.format(dtf);

                    SXSSFCell cell_6_11V = row_6V.createCell(iV);
                    cell_6_11V.setCellStyle(tableHeaderStyle);
                    cell_6_11V.setCellValue(timestamp);

                    SXSSFCell cell_7_11V = row_7V.createCell(iV);
                    cell_7_11V.setCellStyle(tableHeaderStyle);
                    cell_7_11V.setCellValue("Объемный расход воды на нужды горячего водоснабжения");


                    SXSSFCell cell_8_11V = row_8V.createCell(iV);
                    cell_8_11V.setCellStyle(tableHeaderStyle);
                    cell_8_11V.setCellValue(repType.getType());

                    SXSSFCell cell_9_11V = row_9V.createCell(iV);
                    cell_9_11V.setCellStyle(tableHeaderStyle);
                    cell_9_11V.setCellValue("М.куб.");

                    iV++;
                }

                LOGGER.log(Level.INFO, "Report head created {0}", repId);
                // Отлично. Заголовок и шапка сделаны. Идем по таблице, создаем и заполняем ячейки.

                fillSheetV(wb, repId, begRowV, dateListV.size(), dsR, dsRW, repType);

                sh.createFreezePane(9, 9);


                LOGGER.log(Level.INFO, "Report body created {0}", repId);

                break;
            case ("pт"):
            case ("pто"):
            case ("pц"):
            case ("pцо"):
            case ("pг"):
            case ("pго"):
                int begRowP = 9;  // строка в екселе, с которой начинается собственно отчет.
//                long colsP = 0;
                List<LocalDateTime> dateListP = new ArrayList<>();
                LocalDateTime localDateTempP = repType.getBeg();

                dateListP.add(localDateTempP);
                //-- Считаем количество дней между датами

                int colsP = fillDateList(repType, dateListP, localDateTempP);

                //Приступим к основному отчету
                // Устанавливаем ширины колонок. В конце мероприятия
                setColumnWidth(sh);
                sh.setColumnWidth(6, 4695);
                sh.setColumnWidth(7, 4695);

                for (int i = 1; i <= colsP; i++) {
                    sh.setColumnWidth(7+i, 4669);
                }

                // Прекрасно, Заголовок сделали. Готовим шапку.
                SXSSFRow row_6P = sh.createRow(5);
                row_6P.setHeight((short) 350);
                SXSSFRow row_7P = sh.createRow(6);
                row_7P.setHeight((short) 2720);
                SXSSFRow row_8P = sh.createRow(7);
                row_8P.setHeight((short) 350);
                SXSSFRow row_9P = sh.createRow(8);
                row_9P.setHeight((short) 350);


                SXSSFCell cell_6_1P = row_6P.createCell(0);
                cell_6_1P.setCellStyle(tableHeaderStyle);
                cell_6_1P.setCellValue("№ п/п");

                SXSSFCell cell_6_2P = row_6P.createCell(1);
                cell_6_2P.setCellStyle(tableHeaderStyle);
                cell_6_2P.setCellValue("Объект");

                SXSSFCell cell_6_3P = row_6P.createCell(2);
                cell_6_3P.setCellStyle(tableHeaderStyle);
                cell_6_3P.setCellValue("Филиал");

                SXSSFCell cell_6_4P = row_6P.createCell(3);
                cell_6_4P.setCellStyle(tableHeaderStyle);
                cell_6_4P.setCellValue("Предприятие");

                SXSSFCell cell_6_5P = row_6P.createCell(4);
                cell_6_5P.setCellStyle(tableHeaderStyle);
                cell_6_5P.setCellValue("Адрес ЦТП");

                SXSSFCell cell_6_6P = row_6P.createCell(5);
                cell_6_6P.setCellStyle(tableHeaderStyle);
                cell_6_6P.setCellValue("Зона");

                createMergeForTableHeader(sh);

                SXSSFCell cell_6_7P = row_6P.createCell(6);
                cell_6_7P.setCellStyle(tableHeaderStyle);
                cell_6_7P.setCellValue("Нормативный диапазон (нормативная уставка)");
                CellRangeAddress headerNormRangeP = new CellRangeAddress(5, 6, 6, 7);
                sh.addMergedRegion(headerNormRangeP);
                CellRangeAddress borderForNormRangeP = new CellRangeAddress(5, 6, 6, 7);
                setBorders(borderForNormRangeP, sh);

                SXSSFCell cell_8_7P = row_8P.createCell(6);
                cell_8_7P.setCellStyle(tableHeaderStyle);
                cell_8_7P.setCellValue("MIN");

                SXSSFCell cell_8_8P = row_8P.createCell(7);
                cell_8_8P.setCellStyle(tableHeaderStyle);
                cell_8_8P.setCellValue("MAX");

                SXSSFCell cell_9_7P = row_9P.createCell(6);
                cell_9_7P.setCellStyle(tableHeaderStyle);
                cell_9_7P.setCellValue("МПа");
                SXSSFCell cell_9_8P = row_9P.createCell(7);
                cell_9_8P.setCellStyle(tableHeaderStyle);
                cell_9_8P.setCellValue("МПа");

                // декларируем переменные для шапки

                int iP = 8; //вообще последний неизменный столбец - 6, но считаем с 0, потому 5

                for (LocalDateTime localDateTime: dateListP) {

                    DateTimeFormatter dtf = DateTimeFormatter.ofPattern("dd.MM.yyyy");
                    String timestamp = localDateTime.format(dtf);

                    SXSSFCell cell_6_11P = row_6P.createCell(iP);
                    cell_6_11P.setCellStyle(tableHeaderStyle);
                    cell_6_11P.setCellValue(timestamp);

                    SXSSFCell cell_7_11P = row_7P.createCell(iP);
                    cell_7_11P.setCellStyle(tableHeaderStyle);
                    if (repType.getTypeCode().equals("pт")) {
                        cell_7_11P.setCellValue("Среднесуточное давление сетевой воды в подающем трубопроводе Теплосети");
                    }
                    if (repType.getTypeCode().equals("pто")) {
                        cell_7_11P.setCellValue("Среднесуточное давление сетевой воды в обратном трубопроводе Теплосети");
                    }
                    if (repType.getTypeCode().equals("pц")) {
                        cell_7_11P.setCellValue("Среднесуточное давление в подающем трубопроводе ЦО");
                    }
                    if (repType.getTypeCode().equals("pцо")) {
                        cell_7_11P.setCellValue("Среднесуточное давление в обратном трубопроводе ЦО");
                    }
                    if (repType.getTypeCode().equals("pг")) {
                        cell_7_11P.setCellValue("Среднесуточное давление  горячей воды в подающем трубопроводе ГВС");
                    }
                    if (repType.getTypeCode().equals("pго")) {
                        cell_7_11P.setCellValue("Среднесуточное давление  горячей воды в обратном трубопроводе ГВС");
                    }

                    SXSSFCell cell_8_11P = row_8P.createCell(iP);
                    cell_8_11P.setCellStyle(tableHeaderStyle);
                    cell_8_11P.setCellValue(repType.getType());

                    SXSSFCell cell_9_11P = row_9P.createCell(iP);
                    cell_9_11P.setCellStyle(tableHeaderStyle);
                    cell_9_11P.setCellValue("МПа");

                    iP++;
                }

                LOGGER.log(Level.INFO, "Report head created {0}", repId);
                // Отлично. Заголовок и шапка сделаны. Идем по таблице, создаем и заполняем ячейки.

                fillSheetP(wb, repId, begRowP, dateListP.size(), dsR, dsRW, repType);

                sh.createFreezePane(8, 9);

                LOGGER.log(Level.INFO, "Report body created {0}", repId);

                break;
            case ("Gт"):

                int begRowG = 10;  // строка в екселе, с которой начинается собственно отчет.
                int colsG = 0;
                List<LocalDateTime> dateListG = new ArrayList<>();
                LocalDateTime localDateTempG = repType.getBeg();

                dateListG.add(localDateTempG);
                colsG = fillDateList(repType, dateListG, localDateTempG);

                if (repType.getInterval().equals("D")) {
                    colsG = colsG*24;
                }

                //Приступим к основному отчету
                // Устанавливаем ширины колонок. В конце мероприятия
                sh.setColumnWidth(0, 9 * 256);
                sh.setColumnWidth(1, 21 * 256);
                sh.setColumnWidth(2, 2944);
                sh.setColumnWidth(3, 4000);
                sh.setColumnWidth(4, 10956);

                if (repType.getInterval().equals("D")) {
                    for (int i = 1; i <= colsG; i++) {
                        sh.setColumnWidth(4+i, 2944);
                    }
                } else {
                    for (int i = 1; i <= colsG; i++) {
                        sh.setColumnWidth(4 + i, 5400);
                    }
                }

                // Прекрасно, Заголовок сделали. Готовим шапку.
                SXSSFRow row_7G = sh.createRow(6);
                row_7G.setHeight((short) 350);

                SXSSFCell cell_7_1G = row_7G.createCell(0);
                cell_7_1G.setCellStyle(tableHeaderStyle);
                cell_7_1G.setCellValue("№ п/п");

                SXSSFCell cell_7_2G = row_7G.createCell(1);
                cell_7_2G.setCellStyle(tableHeaderStyle);
                cell_7_2G.setCellValue("Объект");

                SXSSFCell cell_7_3G = row_7G.createCell(2);
                cell_7_3G.setCellStyle(tableHeaderStyle);
                cell_7_3G.setCellValue("Филиал");

                SXSSFCell cell_7_4G = row_7G.createCell(3);
                cell_7_4G.setCellStyle(tableHeaderStyle);
                cell_7_4G.setCellValue("Предприятие");

                SXSSFCell cell_7_5G = row_7G.createCell(4);
                cell_7_5G.setCellStyle(tableHeaderStyle);
                cell_7_5G.setCellValue("Адрес ЦТП");

                // декларируем переменные для шапки

                if (repType.getInterval().equals("D")) {
                    SXSSFRow row_8 = sh.createRow(7);
                    SXSSFRow row_9 = sh.createRow(8);
                    SXSSFRow row_10 = sh.createRow(9);
                    SXSSFRow row_11 = sh.createRow(10);
                    row_9.setHeight((short) 2720);


                    begRowG++;
                    CellRangeAddress num = new CellRangeAddress(6, 10, 0, 0);
                    sh.addMergedRegion(num);
                    CellRangeAddress borderForNum = new CellRangeAddress(6, 10, 0, 0);
                    RegionUtil.setBorderBottom(BorderStyle.THICK, borderForNum, sh);
                    RegionUtil.setBorderTop(BorderStyle.THICK, borderForNum, sh);
                    RegionUtil.setBorderLeft(BorderStyle.THICK, borderForNum, sh);
                    RegionUtil.setBorderRight(BorderStyle.THICK, borderForNum, sh);

                    CellRangeAddress objPar = new CellRangeAddress(6, 10, 1, 1);
                    sh.addMergedRegion(objPar);
                    CellRangeAddress borderForObjPar = new CellRangeAddress(6, 10, 1, 1);
                    RegionUtil.setBorderBottom(BorderStyle.THICK, borderForObjPar, sh);
                    RegionUtil.setBorderTop(BorderStyle.THICK, borderForObjPar, sh);
                    RegionUtil.setBorderLeft(BorderStyle.THICK, borderForObjPar, sh);
                    RegionUtil.setBorderRight(BorderStyle.THICK, borderForObjPar, sh);

                    CellRangeAddress techProc = new CellRangeAddress(6, 10, 2, 2);
                    sh.addMergedRegion(techProc);
                    CellRangeAddress borderForTechProc = new CellRangeAddress(6, 10, 2, 2);
                    RegionUtil.setBorderBottom(BorderStyle.THICK, borderForTechProc, sh);
                    RegionUtil.setBorderTop(BorderStyle.THICK, borderForTechProc, sh);
                    RegionUtil.setBorderLeft(BorderStyle.THICK, borderForTechProc, sh);
                    RegionUtil.setBorderRight(BorderStyle.THICK, borderForTechProc, sh);

                    CellRangeAddress unit = new CellRangeAddress(6, 10, 3, 3);
                    sh.addMergedRegion(unit);
                    CellRangeAddress borderForUnit = new CellRangeAddress(6, 10, 3, 3);
                    RegionUtil.setBorderBottom(BorderStyle.THICK, borderForUnit, sh);
                    RegionUtil.setBorderTop(BorderStyle.THICK, borderForUnit, sh);
                    RegionUtil.setBorderLeft(BorderStyle.THICK, borderForUnit, sh);
                    RegionUtil.setBorderRight(BorderStyle.THICK, borderForUnit, sh);

                    CellRangeAddress total = new CellRangeAddress(6, 10, 4, 4);
                    sh.addMergedRegion(total);
                    CellRangeAddress borderForTotal = new CellRangeAddress(6, 10, 4, 4);
                    RegionUtil.setBorderBottom(BorderStyle.THICK, borderForTotal, sh);
                    RegionUtil.setBorderTop(BorderStyle.THICK, borderForTotal, sh);
                    RegionUtil.setBorderLeft(BorderStyle.THICK, borderForTotal, sh);
                    RegionUtil.setBorderRight(BorderStyle.THICK, borderForTotal, sh);

                    for (int i = 0; i < dateListG.size(); i++) {
                        SXSSFCell dateCell = row_7G.createCell(i*24 + 5);
                        dateCell.setCellStyle(tableHeaderStyle);
                        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd.MM.yyyy");
                        String curDateString = dateListG.get(i).format(formatter);
                        dateCell.setCellValue(curDateString);
                        CellRangeAddress date = new CellRangeAddress(6, 6, i*24+5, i*24+28);
                        sh.addMergedRegion(date);
                        CellRangeAddress borderForDate = new CellRangeAddress(6, 6, i*24+5, i*24+28);
                        RegionUtil.setBorderBottom(BorderStyle.THICK, borderForDate, sh);
                        RegionUtil.setBorderTop(BorderStyle.THICK, borderForDate, sh);
                        RegionUtil.setBorderLeft(BorderStyle.THICK, borderForDate, sh);
                        RegionUtil.setBorderRight(BorderStyle.THICK, borderForDate, sh);

                        for (int j = 0; j < 24; j++) {
                            SXSSFCell hourCell = row_8.createCell(i*24 + 5 + j);
                            sh.setColumnWidth(i*24 + 5 + j, 4695);
                            hourCell.setCellStyle(tableHeaderStyle);
                            if (j<9) {
                                hourCell.setCellValue("0" + (j + 1) + " ч.");
                            } else if (j == 23) {
                                hourCell.setCellValue("00 ч.");
                            } else {
                                hourCell.setCellValue((j+1) + " ч.");
                            }
                            SXSSFCell nameCell = row_9.createCell(i*24 + 5 + j);
                            nameCell.setCellStyle(tableHeaderStyle);
                            nameCell.setCellValue("Массовый расход сетевой воды на тепловом вводе");

                            SXSSFCell algNameCell = row_10.createCell(i*24 + 5 + j);
                            algNameCell.setCellStyle(tableHeaderStyle);
                            algNameCell.setCellValue("G1 - G2");

                            SXSSFCell unitsCell = row_11.createCell(i*24 + 5 + j);
                            unitsCell.setCellStyle(tableHeaderStyle);
                            unitsCell.setCellValue("тонн");
                        }
                    }
                } else{
                    SXSSFRow row_8 = sh.createRow(7);
                    SXSSFRow row_9 = sh.createRow(8);
                    SXSSFRow row_10 = sh.createRow(9);
                    row_8.setHeight((short) 2720);

                    CellRangeAddress num = new CellRangeAddress(6, 9, 0, 0);
                    sh.addMergedRegion(num);
                    CellRangeAddress borderForNum = new CellRangeAddress(6, 9, 0, 0);
                    RegionUtil.setBorderBottom(BorderStyle.THICK, borderForNum, sh);
                    RegionUtil.setBorderTop(BorderStyle.THICK, borderForNum, sh);
                    RegionUtil.setBorderLeft(BorderStyle.THICK, borderForNum, sh);
                    RegionUtil.setBorderRight(BorderStyle.THICK, borderForNum, sh);

                    CellRangeAddress objPar = new CellRangeAddress(6, 9, 1, 1);
                    sh.addMergedRegion(objPar);
                    CellRangeAddress borderForObjPar = new CellRangeAddress(6, 9, 1, 1);
                    RegionUtil.setBorderBottom(BorderStyle.THICK, borderForObjPar, sh);
                    RegionUtil.setBorderTop(BorderStyle.THICK, borderForObjPar, sh);
                    RegionUtil.setBorderLeft(BorderStyle.THICK, borderForObjPar, sh);
                    RegionUtil.setBorderRight(BorderStyle.THICK, borderForObjPar, sh);

                    CellRangeAddress techProc = new CellRangeAddress(6, 9, 2, 2);
                    sh.addMergedRegion(techProc);
                    CellRangeAddress borderForTechProc = new CellRangeAddress(6, 9, 2, 2);
                    RegionUtil.setBorderBottom(BorderStyle.THICK, borderForTechProc, sh);
                    RegionUtil.setBorderTop(BorderStyle.THICK, borderForTechProc, sh);
                    RegionUtil.setBorderLeft(BorderStyle.THICK, borderForTechProc, sh);
                    RegionUtil.setBorderRight(BorderStyle.THICK, borderForTechProc, sh);

                    CellRangeAddress unit = new CellRangeAddress(6, 9, 3, 3);
                    sh.addMergedRegion(unit);
                    CellRangeAddress borderForUnit = new CellRangeAddress(6, 9, 3, 3);
                    RegionUtil.setBorderBottom(BorderStyle.THICK, borderForUnit, sh);
                    RegionUtil.setBorderTop(BorderStyle.THICK, borderForUnit, sh);
                    RegionUtil.setBorderLeft(BorderStyle.THICK, borderForUnit, sh);
                    RegionUtil.setBorderRight(BorderStyle.THICK, borderForUnit, sh);

                    CellRangeAddress total = new CellRangeAddress(6, 9, 4, 4);
                    sh.addMergedRegion(total);
                    CellRangeAddress borderForTotal = new CellRangeAddress(6, 9, 4, 4);
                    RegionUtil.setBorderBottom(BorderStyle.THICK, borderForTotal, sh);
                    RegionUtil.setBorderTop(BorderStyle.THICK, borderForTotal, sh);
                    RegionUtil.setBorderLeft(BorderStyle.THICK, borderForTotal, sh);
                    RegionUtil.setBorderRight(BorderStyle.THICK, borderForTotal, sh);

                    int i = 0;
                    for (LocalDateTime curDateLDT : dateListG) {
                        String curDateS = String.valueOf(curDateLDT);
                        sh.setColumnWidth(i + 5, 4695);

                        curDateS = curDateS.replace('T', ' ');
                        SXSSFCell dateCell = row_7G.createCell(i + 5);
                        dateCell.setCellStyle(tableHeaderStyle);
                        curDateS = curDateS.substring(8, 10) + "." + curDateS.substring(5, 7);
                        dateCell.setCellValue(curDateS);

                        SXSSFCell nameCell = row_8.createCell(i + 5);
                        nameCell.setCellStyle(tableHeaderStyle);
                        nameCell.setCellValue("Массовый расход сетевой воды на тепловом вводе");

                        SXSSFCell algNameCell = row_9.createCell(i + 5);
                        algNameCell.setCellStyle(tableHeaderStyle);
                        algNameCell.setCellValue("G1 - G2");

                        SXSSFCell unitsCell = row_10.createCell(i + 5);
                        unitsCell.setCellStyle(tableHeaderStyle);
                        unitsCell.setCellValue("тонн");

                        i++;
                    }
                }

                LOGGER.log(Level.INFO, "Report head created {0}", repId);

                // Отлично. Заголовок и шапка сделаны. Идем по таблице, создаем и заполняем ячейки.

                fillSheetG(wb, repId, begRowG, colsG, dsR, dsRW, repType);

                if (repType.getInterval().equals("D")) {
                    sh.createFreezePane(5, 11);
                } else {
                    sh.createFreezePane(5, 10);
                }

                LOGGER.log(Level.INFO, "Report body created {0}", repId);

                break;

            default:
                break;
        }

        return wb;
    }

    private int fillDateList(RepType repType, List<LocalDateTime> dateList, LocalDateTime localDateTemp) {
        int cols = 0;
        for (;;) {
            if (cols == 0 || repType.getEnd().isAfter(localDateTemp)) {
                if (cols != 0) {
                    localDateTemp = localDateTemp.plusDays(1);
                    dateList.add(localDateTemp);
                }
                cols++;
            } else {
                break;
            }
        }
        return cols;
    }

    private void setColumnWidth(SXSSFSheet sh) {
        sh.setColumnWidth(0, 9 * 256);
        sh.setColumnWidth(1, 21 * 256);
        sh.setColumnWidth(2, 2944);
        sh.setColumnWidth(3, 4000);
        sh.setColumnWidth(4, 10956);
        sh.setColumnWidth(5, 9 * 256);

    }

    private void createMergeForTableHeader(SXSSFSheet sh) {

        CellRangeAddress headerNum = new CellRangeAddress(5, 8, 0, 0);
        sh.addMergedRegion(headerNum);
        CellRangeAddress borderForNum = new CellRangeAddress(5, 8, 0, 0);
        RegionUtil.setBorderBottom(BorderStyle.THICK, borderForNum, sh);
        RegionUtil.setBorderTop(BorderStyle.THICK, borderForNum, sh);
        RegionUtil.setBorderLeft(BorderStyle.THICK, borderForNum, sh);
        RegionUtil.setBorderRight(BorderStyle.THICK, borderForNum, sh);

        CellRangeAddress headerObj = new CellRangeAddress(5, 8, 1, 1);
        sh.addMergedRegion(headerObj);
        CellRangeAddress borderForObj = new CellRangeAddress(5, 8, 1, 1);
        RegionUtil.setBorderBottom(BorderStyle.THICK, borderForObj, sh);
        RegionUtil.setBorderTop(BorderStyle.THICK, borderForObj, sh);
        RegionUtil.setBorderLeft(BorderStyle.THICK, borderForObj, sh);
        RegionUtil.setBorderRight(BorderStyle.THICK, borderForObj, sh);

        CellRangeAddress headerBranch = new CellRangeAddress(5, 8, 2, 2);
        sh.addMergedRegion(headerBranch);
        CellRangeAddress borderForBranch = new CellRangeAddress(5, 8, 2, 2);
        RegionUtil.setBorderBottom(BorderStyle.THICK, borderForBranch, sh);
        RegionUtil.setBorderTop(BorderStyle.THICK, borderForBranch, sh);
        RegionUtil.setBorderLeft(BorderStyle.THICK, borderForBranch, sh);
        RegionUtil.setBorderRight(BorderStyle.THICK, borderForBranch, sh);

        CellRangeAddress headerFacility = new CellRangeAddress(5, 8, 3, 3);
        sh.addMergedRegion(headerFacility);
        CellRangeAddress borderForFacility = new CellRangeAddress(5, 8, 3, 3);
        RegionUtil.setBorderBottom(BorderStyle.THICK, borderForFacility, sh);
        RegionUtil.setBorderTop(BorderStyle.THICK, borderForFacility, sh);
        RegionUtil.setBorderLeft(BorderStyle.THICK, borderForFacility, sh);
        RegionUtil.setBorderRight(BorderStyle.THICK, borderForFacility, sh);

        CellRangeAddress headerAddress = new CellRangeAddress(5, 8, 4, 4);
        sh.addMergedRegion(headerAddress);
        CellRangeAddress borderForAddress = new CellRangeAddress(5, 8, 4, 4);
        RegionUtil.setBorderBottom(BorderStyle.THICK, borderForAddress, sh);
        RegionUtil.setBorderTop(BorderStyle.THICK, borderForAddress, sh);
        RegionUtil.setBorderLeft(BorderStyle.THICK, borderForAddress, sh);
        RegionUtil.setBorderRight(BorderStyle.THICK, borderForAddress, sh);

        CellRangeAddress headerZone = new CellRangeAddress(5, 8, 5, 5);
        sh.addMergedRegion(headerZone);
        CellRangeAddress borderForZone = new CellRangeAddress(5, 8, 5, 5);
        RegionUtil.setBorderBottom(BorderStyle.THICK, borderForZone, sh);
        RegionUtil.setBorderTop(BorderStyle.THICK, borderForZone, sh);
        RegionUtil.setBorderLeft(BorderStyle.THICK, borderForZone, sh);
        RegionUtil.setBorderRight(BorderStyle.THICK, borderForZone, sh);
    }

    private void fillSheetT1 (SXSSFWorkbook wb, int repId, int begRow, int cols, DataSource dsR, DataSource dsRW, RepType repType) throws DecoderException {
        CellStyle cellNoBoldStyle = setCellNoBoldStyle(wb);
        SXSSFSheet sh = wb.getSheetAt(0);

        // Заполняем лист значениями, взятыми из таблицы
        List<ReportObject> objects = reworkObjList(repId, dsR, dsRW, cols, repType);
        Rows = begRow;

        for (ReportObject object: objects) {
                SXSSFRow row = sh.createRow(Rows);
                row.setHeight((short) 350);
                SXSSFCell objNumCell = row.createCell(0);
                objNumCell.setCellValue(object.getNumPP());
                objNumCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell objNameCell = row.createCell(1);
                objNameCell.setCellValue(object.getObjName());
                objNameCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell filialCell = row.createCell(2);
                filialCell.setCellValue(object.getFilial());
                filialCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell predprCell = row.createCell(3);
                predprCell.setCellValue(object.getPredpr());
                predprCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell objAddrCell = row.createCell(4);
                objAddrCell.setCellValue(object.getObjAddress());
                objAddrCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell zoneCell = row.createCell(5);
                zoneCell.setCellValue(object.getValues().get(0).getZone());
                zoneCell.setCellStyle(cellNoBoldStyle);

                int i = 6;
                for (Value value : object.getValues()){

                    SXSSFCell tnv = row.createCell(i);
                    tnv.setCellValue(value.getTnv());
                    tnv.setCellStyle(cellNoBoldStyle);

                    SXSSFCell tnvGmc = row.createCell(i+1);
                    tnvGmc.setCellValue(value.getTnvGmc());
                    tnvGmc.setCellStyle(cellNoBoldStyle);

                    SXSSFCell min = row.createCell(i+2);
                    min.setCellValue(value.getMin());
                    min.setCellStyle(cellNoBoldStyle);

                    SXSSFCell max = row.createCell(i+3);
                    max.setCellValue(value.getMax());
                    max.setCellStyle(cellNoBoldStyle);

                    SXSSFCell parValue = row.createCell(i+4);
                    parValue.setCellValue(value.getParValue());
                    if (value.getColor() != null) {
                        if (colors.containsKey(value.getColor())) {
                            parValue.setCellStyle(colors.get(value.getColor()));
                        } else {
                            CellStyle cellColoredStyle = setCellNoBoldStyle(wb);
                            String rgbS = value.getColor();
                            byte [] rgbB = Hex.decodeHex(rgbS);
                            XSSFColor color = new XSSFColor(rgbB, null);
                            cellColoredStyle.setFillForegroundColor(color);
                            cellColoredStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                            colors.put(value.getColor(), cellColoredStyle);
                            parValue.setCellStyle(cellColoredStyle);
                        }
                    } else {
                        parValue.setCellStyle(cellNoBoldStyle);
                    }
                    i = i + 5;
                }
                Rows++;
        }
    }

    private void fillSheetT7 (SXSSFWorkbook wb, int repId, int begRow, int cols, DataSource dsR, DataSource dsRW, RepType repType) throws DecoderException {
        CellStyle cellNoBoldStyle = setCellNoBoldStyle(wb);
        SXSSFSheet sh = wb.getSheetAt(0);

        // Заполняем лист значениями, взятыми из таблицы
        List<ReportObject> objects = reworkObjList(repId, dsR, dsRW, cols, repType);

        Rows = begRow;


        for (ReportObject object: objects) {
                SXSSFRow row = sh.createRow(Rows);
                row.setHeight((short) 350);
                SXSSFCell objNumCell = row.createCell(0);
                objNumCell.setCellValue(object.getNumPP());
                objNumCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell objNameCell = row.createCell(1);
                objNameCell.setCellValue(object.getObjName());
                objNameCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell filialCell = row.createCell(2);
                filialCell.setCellValue(object.getFilial());
                filialCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell predprCell = row.createCell(3);
                predprCell.setCellValue(object.getPredpr());
                predprCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell objAddrCell = row.createCell(4);
                objAddrCell.setCellValue(object.getObjAddress());
                objAddrCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell zoneCell = row.createCell(5);
                zoneCell.setCellValue(object.getValues().get(0).getZone());
                zoneCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell tnv = row.createCell(6);
                tnv.setCellValue(object.getValues().get(0).getTnv());
                tnv.setCellStyle(cellNoBoldStyle);

                SXSSFCell tnvGmc = row.createCell(7);
                tnvGmc.setCellValue(object.getValues().get(0).getTnvGmc());
                tnvGmc.setCellStyle(cellNoBoldStyle);

                SXSSFCell min = row.createCell(8);
                min.setCellValue(object.getValues().get(0).getMin());
                min.setCellStyle(cellNoBoldStyle);

                SXSSFCell max = row.createCell(9);
                max.setCellValue(object.getValues().get(0).getMax());
                max.setCellStyle(cellNoBoldStyle);

                int i = 10;
                for (Value value : object.getValues()){

                    SXSSFCell parValue = row.createCell(i);
                    parValue.setCellValue(value.getParValue());
                    if (value.getColor() != null) {
                        if (colors.containsKey(value.getColor())) {
                            parValue.setCellStyle(colors.get(value.getColor()));
                        } else {
                            CellStyle cellColoredStyle = setCellNoBoldStyle(wb);
                            String rgbS = value.getColor();
                            byte [] rgbB = Hex.decodeHex(rgbS);
                            XSSFColor color = new XSSFColor(rgbB, null);
                            cellColoredStyle.setFillForegroundColor(color);
                            cellColoredStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                            colors.put(value.getColor(), cellColoredStyle);
                            parValue.setCellStyle(cellColoredStyle);
                        }
                    } else {
                        parValue.setCellStyle(cellNoBoldStyle);
                    }
                    i++;
                }
                Rows++;
        }
    }

    private void fillSheetV (SXSSFWorkbook wb, int repId, int begRow, int cols, DataSource dsR, DataSource dsRW, RepType repType) throws DecoderException {
        CellStyle cellNoBoldStyle = setCellNoBoldStyle(wb);
        SXSSFSheet sh = wb.getSheetAt(0);

        // Заполняем лист значениями, взятыми из таблицы
        List<ReportObject> objects = reworkObjList(repId, dsR, dsRW, cols, repType);

        Rows = begRow;

        for (ReportObject object: objects) {
                SXSSFRow row = sh.createRow(Rows);
                row.setHeight((short) 350);
                SXSSFCell objNumCell = row.createCell(0);
                objNumCell.setCellValue(object.getNumPP());
                objNumCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell objNameCell = row.createCell(1);
                objNameCell.setCellValue(object.getObjName());
                objNameCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell filialCell = row.createCell(2);
                filialCell.setCellValue(object.getFilial());
                filialCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell predprCell = row.createCell(3);
                predprCell.setCellValue(object.getPredpr());
                predprCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell objAddrCell = row.createCell(4);
                objAddrCell.setCellValue(object.getObjAddress());
                objAddrCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell zoneCell = row.createCell(5);
                zoneCell.setCellValue(object.getValues().get(0).getZone());
                zoneCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell measureCell = row.createCell(6);
                measureCell.setCellValue(object.getValues().get(0).getMeasure());
                measureCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell tnv = row.createCell(7);
                tnv.setCellValue(object.getValues().get(0).getMin());
                tnv.setCellStyle(cellNoBoldStyle);

                SXSSFCell tnvGmc = row.createCell(8);
                tnvGmc.setCellValue(object.getValues().get(0).getMax());
                tnvGmc.setCellStyle(cellNoBoldStyle);

                int i = 9;
                for (Value value : object.getValues()){

                    SXSSFCell parValue = row.createCell(i);
                    parValue.setCellValue(value.getParValue());
                    if (value.getColor() != null) {
                        if (colors.containsKey(value.getColor())) {
                            parValue.setCellStyle(colors.get(value.getColor()));
                        } else {
                            CellStyle cellColoredStyle = setCellNoBoldStyle(wb);
                            String rgbS = value.getColor();
                            byte [] rgbB = Hex.decodeHex(rgbS);
                            XSSFColor color = new XSSFColor(rgbB, null);
                            cellColoredStyle.setFillForegroundColor(color);
                            cellColoredStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                            colors.put(value.getColor(), cellColoredStyle);
                            parValue.setCellStyle(cellColoredStyle);
                        }
                    } else {
                        parValue.setCellStyle(cellNoBoldStyle);
                    }
                    i++;
                }
                Rows++;
        }
    }

    private void fillSheetP (SXSSFWorkbook wb, int repId, int begRow, int cols, DataSource dsR, DataSource dsRW, RepType repType) throws DecoderException {
        CellStyle cellNoBoldStyle = setCellNoBoldStyle(wb);
        SXSSFSheet sh = wb.getSheetAt(0);

        // Заполняем лист значениями, взятыми из таблицы
        List<ReportObject> objects = reworkObjList(repId, dsR, dsRW, cols, repType);

        Rows = begRow;


        for (ReportObject object: objects) {
                SXSSFRow row = sh.createRow(Rows);
                row.setHeight((short) 350);
                SXSSFCell objNumCell = row.createCell(0);
                objNumCell.setCellValue(object.getNumPP());
                objNumCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell objNameCell = row.createCell(1);
                objNameCell.setCellValue(object.getObjName());
                objNameCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell filialCell = row.createCell(2);
                filialCell.setCellValue(object.getFilial());
                filialCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell predprCell = row.createCell(3);
                predprCell.setCellValue(object.getPredpr());
                predprCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell objAddrCell = row.createCell(4);
                objAddrCell.setCellValue(object.getObjAddress());
                objAddrCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell zoneCell = row.createCell(5);
                zoneCell.setCellValue(object.getValues().get(0).getZone());
                zoneCell.setCellStyle(cellNoBoldStyle);
                SXSSFCell tnv = row.createCell(6);
                tnv.setCellValue(object.getValues().get(0).getMin());
                tnv.setCellStyle(cellNoBoldStyle);

                SXSSFCell tnvGmc = row.createCell(7);
                tnvGmc.setCellValue(object.getValues().get(0).getMax());
                tnvGmc.setCellStyle(cellNoBoldStyle);

                int i = 8;
                for (Value value : object.getValues()){


                    SXSSFCell parValue = row.createCell(i);
                    parValue.setCellValue(value.getParValue());
                    if (value.getColor() != null) {
                        if (colors.containsKey(value.getColor())) {
                            parValue.setCellStyle(colors.get(value.getColor()));
                        } else {
                            CellStyle cellColoredStyle = setCellNoBoldStyle(wb);
                            String rgbS = value.getColor();
                            byte [] rgbB = Hex.decodeHex(rgbS);
                            XSSFColor color = new XSSFColor(rgbB, null);
                            cellColoredStyle.setFillForegroundColor(color);
                            cellColoredStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                            colors.put(value.getColor(), cellColoredStyle);
                            parValue.setCellStyle(cellColoredStyle);
                        }
                    } else {
                        parValue.setCellStyle(cellNoBoldStyle);
                    }
                    i++;
                }
                Rows++;
        }
    }

    private void fillSheetG (SXSSFWorkbook wb, int repId, int begRow, int cols, DataSource dsR, DataSource dsRW, RepType repType) throws DecoderException {
        CellStyle cellNoBoldStyle = setCellNoBoldStyle(wb);
        SXSSFSheet sh = wb.getSheetAt(0);

        // Заполняем лист значениями, взятыми из таблицы
        List<ReportObject> objects = reworkObjList(repId, dsR, dsRW, cols, repType);

        Rows = begRow;

        for (ReportObject object: objects) {
            SXSSFRow row = sh.createRow(Rows);
            row.setHeight((short) 350);
            SXSSFCell objNumCell = row.createCell(0);
            objNumCell.setCellValue(object.getNumPP());
            objNumCell.setCellStyle(cellNoBoldStyle);
            SXSSFCell objNameCell = row.createCell(1);
            objNameCell.setCellValue(object.getObjName());
            objNameCell.setCellStyle(cellNoBoldStyle);
            SXSSFCell filialCell = row.createCell(2);
            filialCell.setCellValue(object.getFilial());
            filialCell.setCellStyle(cellNoBoldStyle);
            SXSSFCell predprCell = row.createCell(3);
            predprCell.setCellValue(object.getPredpr());
            predprCell.setCellStyle(cellNoBoldStyle);
            SXSSFCell objAddrCell = row.createCell(4);
            objAddrCell.setCellValue(object.getObjAddress());
            objAddrCell.setCellStyle(cellNoBoldStyle);

            int i = 5;
            for (Value value : object.getValues()){

                SXSSFCell parValue = row.createCell(i);
                parValue.setCellValue(value.getParValue());
                if (value.getColor() != null) {
                    if (colors.containsKey(value.getColor())) {
                        parValue.setCellStyle(colors.get(value.getColor()));
                    } else {
                        CellStyle cellColoredStyle = setCellNoBoldStyle(wb);
                        String rgbS = value.getColor();
                        byte [] rgbB = Hex.decodeHex(rgbS);
                        XSSFColor color = new XSSFColor(rgbB, null);
                        cellColoredStyle.setFillForegroundColor(color);
                        cellColoredStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                        colors.put(value.getColor(), cellColoredStyle);
                        parValue.setCellStyle(cellColoredStyle);
                    }
                } else {
                    parValue.setCellStyle(cellNoBoldStyle);
                }
                i++;
            }
            Rows++;
        }
    }

    public List<ReportObject> reworkObjList(int repId, DataSource dsR, DataSource dsRW, int cols, RepType repType) {
        List<ReportObject> result = new ArrayList<>();
        List<ReportObject> tempObj = loadObjects(repId, dsR);
        BigDecimal percentage = new BigDecimal(0);
        double size = tempObj.size();
        double iterationPercent = 100/size;
        BigDecimal iterationPercentBD = new BigDecimal(iterationPercent);

        for (ReportObject object : tempObj) {
            if (!interrupted(repId, dsR).equals("Q")) {
                switch (repType.getTypeCode()) {
                    case ("Tт"):
                    case ("Tто"):
                    case ("Tц"):
                    case ("Tцо"):
                        object.setValues(loadValuesT1(repId, object.getObjId(), dsR));
                        break;
                    case ("Tг"):
                    case ("Tго"):
                        object.setValues(loadValuesT7(repId, object.getObjId(), dsR));
                        break;
                    case ("Qгп"):
                        object.setValues(loadValuesV(repId, object.getObjId(), dsR));
                        break;
                    case ("pт"):
                    case ("pто"):
                    case ("pц"):
                    case ("pцо"):
                    case ("pг"):
                    case ("pго"):
                        object.setValues(loadValuesP(repId, object.getObjId(), dsR));
                        break;
                    case ("Gт"):
                        object.setValues(loadValuesG(repId, object.getObjId(), dsR));
                        break;

                }
                if (object.getValues().isEmpty()) {
                    List<Value> values = new ArrayList<>();
                    for (int i = 0; i < cols; i++) {
                        values.add(new Value("", "", "", "", "", "", "ffffff"));
                    }
                    object.setValues(values);
                }
                    if (object.getValues().size() == cols) {
                        result.add(object);
                    } else {
                        for (int i = 0; i < (object.getValues().size()/cols); i++) {
                            ReportObject tempObject = new ReportObject(object.getNumPP(), object.getObjId(), object.getObjName(), object.getFilial(),
                                    object.getPredpr(), object.getObjAddress());
                            List<Value> tempValue = new ArrayList<>();
                            for (int k = 0; k < cols; k++) {
                                tempValue.add(object.getValues().get(k + i*cols));
                            }
                            tempObject.setValues(tempValue);
                            result.add(tempObject);
                        }
                    }
            } else {
                break;
            }
            percentage = percentage.add(iterationPercentBD).setScale(3, RoundingMode.DOWN);
            percent(repId, percentage, dsRW);
        }
        return result;
    }

    public RepType loadRepType(int repId, DataSource ds) {
        RepType result = new RepType();
        try (Connection connect = ds.getConnection();
             PreparedStatement stm = connect.prepareStatement(LOAD_REP_TYPE)) {
            stm.setInt(1, repId);
            ResultSet res = stm.executeQuery();
            if (res.next()) {
                result.setBeg(res.getTimestamp("beg_date").toLocalDateTime());
                result.setEnd(res.getTimestamp("end_date").toLocalDateTime());
                result.setType(res.getString("rep_type"));
                result.setInterval(res.getString("interval"));
                result.setTypeCode(res.getString("par_code"));
                return result;
            }

        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "error load Rep Type", e);
        }
        return result;
    }

    //получаем список объектов в
    public List<ReportObject> loadObjects(int repId, DataSource ds) {
        List<ReportObject> result = new ArrayList<>();
        try (Connection connect = ds.getConnection();
             PreparedStatement stm = connect.prepareStatement(LOAD_OBJECT)) {
            stm.setInt(1, repId);
            ResultSet res = stm.executeQuery();
            while (res.next()) {
                ReportObject item = new ReportObject(res.getInt("n_id"), res.getInt("obj_id"),
                        res.getString("obj_name"), res.getString("filial"),
                        res.getString("predpr"), res.getString("obj_address"));
                result.add(item);
            }
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "error load Object", e);
        }
        return result;
    }

    public List<Value> loadValuesT1(int repId, int objId, DataSource ds) {
        List<Value> result = new ArrayList<>();

        try (Connection connect = ds.getConnection();
             PreparedStatement stm = connect.prepareStatement(LOAD_VALUE_T1_TYPE)) {
            stm.setInt(1, repId);
            stm.setInt(2, objId);
            ResultSet res = stm.executeQuery();
            while (res.next()) {
                result.add(new Value(res.getString("zone"), res.getString("tnv"), res.getString("tnv_gmc"),
                        res.getString("min"), res.getString("max"), res.getString("par_value"),
                        res.getString("color")));
            }
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "error load Values", e);
        }

        return result;
    }

    public List<Value> loadValuesT7(int repId, int objId, DataSource ds) {
        List<Value> result = new ArrayList<>();

        try (Connection connect = ds.getConnection();
             PreparedStatement stm = connect.prepareStatement(LOAD_VALUE_T7_TYPE)) {
            stm.setInt(1, repId);
            stm.setInt(2, objId);
            ResultSet res = stm.executeQuery();
            while (res.next()) {
                result.add(new Value(res.getString("zone"), res.getString("t_min"), res.getString("t_max"),
                        res.getString("a_min"), res.getString("a_max"), res.getString("par_value"),
                        res.getString("color")));
            }
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "error load Values", e);
        }

        return result;
    }

    public List<Value> loadValuesV(int repId, int objId, DataSource ds) {
        List<Value> result = new ArrayList<>();

        try (Connection connect = ds.getConnection();
             PreparedStatement stm = connect.prepareStatement(LOAD_VALUE_V_TYPE)) {
            stm.setInt(1, repId);
            stm.setInt(2, objId);
            ResultSet res = stm.executeQuery();
            while (res.next()) {
                result.add(new Value(res.getString("zone"), res.getString("min"),
                        res.getString("max"), res.getString("par_value"),
                        res.getString("color"), res.getString("measure")));
            }
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "error load Values", e);
        }

        return result;
    }

    public List<Value> loadValuesP(int repId, int objId, DataSource ds) {
        List<Value> result = new ArrayList<>();

        try (Connection connect = ds.getConnection();
             PreparedStatement stm = connect.prepareStatement(LOAD_VALUE_P_TYPE)) {
            stm.setInt(1, repId);
            stm.setInt(2, objId);
            ResultSet res = stm.executeQuery();
            while (res.next()) {
                result.add(new Value(res.getString("zone"), res.getString("min"),
                        res.getString("max"), res.getString("par_value"),
                        res.getString("color")));
            }
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "error load Values", e);
        }

        return result;
    }

    public List<Value> loadValuesG(int repId, int objId, DataSource ds) {
        List<Value> result = new ArrayList<>();

        try (Connection connect = ds.getConnection();
             PreparedStatement stm = connect.prepareStatement(LOAD_VALUE_G_TYPE)) {
            stm.setInt(1, repId);
            stm.setInt(2, objId);
            ResultSet res = stm.executeQuery();
            while (res.next()) {
                result.add(new Value(res.getString("par_value"),
                        res.getString("color")));
            }
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "error load Values", e);
        }

        return result;
    }

    //проверяем не был ли отчет прерван
    public String interrupted(int repId, DataSource ds) {
        try (Connection connection = ds.getConnection();
             PreparedStatement stm = connection.prepareStatement(INTERRUPTED)) {
            stm.setInt(1, repId);
            ResultSet res = stm.executeQuery();
            if (res.next() && (res.getString(1) != null)) {
                return res.getString(1);
            }
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "error load interrupt status ", e);
        }
        return null;
//        return "Q";
    }

    // записываем процент выполнения в таблицу
    public void percent(int repId, BigDecimal percent, DataSource ds) {
        try (Connection connection = ds.getConnection();
             PreparedStatement stm = connection.prepareStatement(PERCENT)) {
            stm.setInt(1, repId);
            stm.setBigDecimal(2, percent);
            stm.executeUpdate();
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "update percent error ", e);
        }
    }

    public void saveReportIntoTable(SXSSFWorkbook wb, int repId, DataSource dsRW) {
            if  (!interrupted(repId, dsRW).equals("Q")) {
                addBlob(wb, repId, dsRW);
                percent(repId, BigDecimal.valueOf(100), dsRW);
                setStatus(repId, dsRW);
            } else {
                delBlob(repId, dsRW);
            }
    }

    //удаляем блоб
    public void delBlob(int repId, DataSource ds) {
        try (Connection connection = ds.getConnection();
             CallableStatement delStmt = connection.prepareCall(DELSQL)) {
            delStmt.setLong(1, repId);
            delStmt.registerOutParameter(2, Types.VARCHAR);
            delStmt.executeUpdate();

        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "Ошибка удаления отчета ", e);
        }
    }

    //добавляеь блоб в таблицу
    public void addBlob(SXSSFWorkbook wb, int repId, DataSource ds) {
        try (Connection connection = ds.getConnection();
             PreparedStatement stm = connection.prepareStatement(SQL)) {

            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try {
                wb.write(bos);
            } catch (IOException e) {
                LOGGER.log(Level.WARNING, "saving blob error ", e);
            }

            stm.setLong(1, repId);
            stm.setBytes(2, bos.toByteArray());

            stm.executeQuery();
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "add blob error ", e);
        }
    }

    //устанавливаем статус
    public void setStatus(int repId, DataSource ds) {
        try (Connection connection = ds.getConnection();
             PreparedStatement stm = connection.prepareStatement(FINSQL)) {
            stm.setInt(1, repId);
            stm.executeUpdate();
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "update Status error ", e);
        }

    }

    public void setHeader(SXSSFSheet sh, CellStyle headerStyle, CellStyle headerStyleNoBold, CellStyle nowStyle, RepType repType) {
        SXSSFRow row_1 = sh.createRow(0);
        row_1.setHeight((short) 435);
        SXSSFCell cell_1_1 = row_1.createCell(0);
        cell_1_1.setCellValue("ПАО \"МОЭК\": АС \"ТЕКОН - Диспетчеризация\"");

        CellRangeAddress title = new CellRangeAddress(0, 0, 0, 5);
        sh.addMergedRegion(title);
        cell_1_1.setCellStyle(headerStyle);

        SXSSFRow row_2 = sh.createRow(1);
        row_2.setHeight((short) 435);
        SXSSFCell cell_2_1 = row_2.createCell(0);
        cell_2_1.setCellValue("Анализ фактической работы ЦТП по показателю: " + repType.getType());
        CellRangeAddress formName = new CellRangeAddress(1, 1, 0, 5);
        sh.addMergedRegion(formName);
        cell_2_1.setCellStyle(headerStyle);

        SXSSFRow row_3 = sh.createRow(2);
        row_3.setHeight((short) 435);
        SXSSFCell cell_3_1 = row_3.createCell(0);
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd.MM.yyyy");
        LocalDateTime begFormatted = repType.getBeg();
        String stringBeg = begFormatted.format(formatter);
        LocalDateTime endFormattedT1 = repType.getEnd();
        String stringEnd = endFormattedT1.format(formatter);
        cell_3_1.setCellValue("за период: " + stringBeg + " - " + stringEnd);
        cell_3_1.setCellStyle(headerStyleNoBold);
        CellRangeAddress period = new CellRangeAddress(2, 2, 0, 5);
        sh.addMergedRegion(period);
        cell_3_1.setCellStyle(headerStyleNoBold);

        // Печатаем отчетов зад общий для всех отчетов
        String now = new SimpleDateFormat("dd.MM.yyyy HH:mm").format(new Date());
        SXSSFRow row_4 = sh.createRow(3);
        row_4.setHeight((short) 435);
        SXSSFCell cell4_1 = row_4.createCell(0);
        cell4_1.setCellStyle(nowStyle);
        cell4_1.setCellValue("Отчет сформирован  " + now);
        CellRangeAddress nowDone = new CellRangeAddress(3, 3, 0, 5);
        sh.addMergedRegion(nowDone);
    }

        public void setBorders(CellRangeAddress border, SXSSFSheet sh) {
        RegionUtil.setBorderBottom(BorderStyle.THICK, border, sh);
        RegionUtil.setBorderTop(BorderStyle.THICK, border, sh);
        RegionUtil.setBorderLeft(BorderStyle.THICK, border, sh);
        RegionUtil.setBorderRight(BorderStyle.THICK, border, sh);
    }

    ///////////////////////////////////////////  Определение стилей тут

    //  Стиль заголовка жирный
    private  CellStyle setHeaderStyle(SXSSFWorkbook p_wb) {

        CellStyle style = p_wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setWrapText(false);

        Font headerFont = p_wb.createFont();
        headerFont.setBold(true);
        headerFont.setFontName("Times New Roman");
        headerFont.setFontHeightInPoints((short) 16);

        style.setFont(headerFont);

        return style;
    }

    //  Стиль заголовка не жирный
    private  CellStyle setHeaderStyleNoBold(SXSSFWorkbook p_wb) {

        CellStyle style = p_wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setWrapText(false);

        Font headerFontNoBold = p_wb.createFont();
        headerFontNoBold.setBold(false);
        headerFontNoBold.setFontName("Times New Roman");
        headerFontNoBold.setFontHeightInPoints((short) 16);

        style.setFont(headerFontNoBold);

        return style;
    }

    //стиль для даты создания отчета
    private  CellStyle setCellNow(SXSSFWorkbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);


        Font nowFont = wb.createFont();
        nowFont.setBold(false);
        nowFont.setFontName("Times New Roman");
        nowFont.setFontHeightInPoints((short) 12);

        style.setFont(nowFont);

        return style;
    }

    //  Стиль шапки таблицы
    private  CellStyle setTableHeaderStyle(SXSSFWorkbook p_wb) {
        CellStyle style = p_wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);

        style.setBorderTop(BorderStyle.THICK);
        style.setBorderLeft(BorderStyle.THICK);
        style.setBorderRight(BorderStyle.THICK);
        style.setBorderBottom(BorderStyle.THICK);

        Font tableHeaderFont = p_wb.createFont();

        tableHeaderFont.setBold(true);
        tableHeaderFont.setFontName("Times New Roman");
        tableHeaderFont.setFontHeightInPoints((short) 12);

        style.setFont(tableHeaderFont);

        return style;
    }

    private  CellStyle setCellNoBoldStyle(SXSSFWorkbook p_wb) {
        CellStyle style = p_wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.TOP);
        style.setWrapText(true);

        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);

        Font cellNoBoldFont = p_wb.createFont();

        cellNoBoldFont.setBold(false);
        cellNoBoldFont.setFontName("Times New Roman");
        cellNoBoldFont.setFontHeightInPoints((short) 11);

        style.setFont(cellNoBoldFont);

        return style;
    }
}
