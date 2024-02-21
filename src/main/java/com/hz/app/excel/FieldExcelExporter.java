package com.hz.app.excel;

import com.google.common.collect.ImmutableMap;
import com.hz.app.excel.model.MyEntity;
import io.swagger.annotations.ApiModelProperty;
import io.swagger.v3.oas.annotations.media.Schema;
import lombok.*;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.time.LocalDateTime;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Excel 工厂类，用于生成 Excel 表格
 *
 * @author hz
 * @since 2024/1/12
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class FieldExcelExporter {

    private static String projectName;

    private static String packageNamePrefix;

    private String modelName;

    private List<MyRow> myRows;

    public static int SHEET_LEVEL = 1;

    public static Deque<Class<?>> CUSTOM_CLASS_LIST = new ArrayDeque<>();
    public static List<List<MyRow>> MY_ROWS_LIST = new ArrayList<>();

    public static Set<Class<?>> CUSTOM_CLASS_SET = new HashSet<>();

    public List<Class<?>> customClassListCopy;

    @Data
    @AllArgsConstructor
    @NoArgsConstructor
    @Builder
    static class MyRow {
        /**
         * 字段名
         */
        private String fieldName;
        /**
         * 字段描述
         */
        private String fieldDesc;
        /**
         * 是否键值
         */
        private boolean isKey;
        /**
         * 字段属性
         */
        private String fieldType;
        /**
         * 备注
         */
        private String remark;
        /**
         * 枚举值
         */
        private String enumValues;
        /**
         * 对象
         */
        private String linkObj;
        /**
         * 同步
         */
        private String syncType;
    }

    private static String toUnderlineCase(String str) {
        Pattern compile = Pattern.compile("[A-Z]");
        Matcher matcher = compile.matcher(str);
        StringBuffer sb = new StringBuffer();
        while (matcher.find()) {
            matcher.appendReplacement(sb, "_" + matcher.group(0).toLowerCase());
        }
        matcher.appendTail(sb);
        if (sb.charAt(0) == '_') {
            sb.deleteCharAt(0);
        }
        return sb.toString();
    }

    /**
     * 初始化 Excel 表格数据
     */
    public void initExcelData(Class<?> clazz, List<MyRow> myRows) {

        if (this.modelName.equals("_init_")) {
            this.modelName = toUnderlineCase(clazz.getSimpleName());
        }

        // 获取 Entity 类的所有字段名和字段类型
        for (Field field : clazz.getDeclaredFields()) {
            MyRow myRow = MyRow.builder()
                    .fieldName(toUnderlineCase(field.getName()))
                    .build();
            ApiModelProperty annotation = field.getAnnotation(ApiModelProperty.class);
            if (annotation != null) {
                myRow.setFieldDesc(annotation.value());
            } else {
                Schema schema = field.getAnnotation(Schema.class);
                if (schema != null) {
                    myRow.setFieldDesc(schema.description());
                }
            }
            if (isCustomClass(field)) {
                Class<?> fieldType = field.getType();
                if (CUSTOM_CLASS_SET.add(fieldType)) {
                    if (List.class.isAssignableFrom(fieldType)) {
                        ParameterizedType genericType = (ParameterizedType) field.getGenericType();
                        fieldType = (Class<?>) genericType.getActualTypeArguments()[0];
                    }
                    CUSTOM_CLASS_LIST.add(fieldType);
                }
                myRow.setFieldType("引用值");
                myRow.setLinkObj(toUnderlineCase(field.getType().getSimpleName()));
                initExcelData(field.getType(), new ArrayList<>());
            } else {
                myRow.setFieldType(getDisplayFormatForFieldType(field.getType().getSimpleName()));
            }

            myRows.add(myRow);
            customClassListCopy = new ArrayList<>(CUSTOM_CLASS_LIST);
        }
    }

    private static boolean isCustomClass(Field field) {
        Class<?> fieldType = field.getType();

        if (List.class.isAssignableFrom(fieldType)) {
            ParameterizedType genericType = (ParameterizedType) field.getGenericType();
            fieldType = (Class<?>) genericType.getActualTypeArguments()[0];
        }

        return fieldType.getPackage().getName().startsWith(packageNamePrefix);
    }


    private static final Map<String, String> DISPLAY_FORMAT_MAP = ImmutableMap.<String, String>builder()
            .put("String", "文本")
            .put("Character", "文本")
            .put("Integer", "数字")
            .put("Byte", "数字")
            .put("Long", "数字")
            .put("Double", "数字")
            .put("Float", "数字")
            .put("BigDecimal", "数字")
            .put("Date", "日期")
            .put("LocalDate", "日期")
            .put("LocalDateTime", "日期")
            .put("LocalTime", "时间")
            .build();

    private static String getDisplayFormatForFieldType(String fieldType) {
        return DISPLAY_FORMAT_MAP.getOrDefault(fieldType, "");
    }

    /**
     * 创建 Excel 工作簿
     *
     * @return Excel 工作簿对象
     */
    public Workbook createWorkbook() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        createSheet(workbook, this.myRows, this.modelName);
        customClassListCopy.forEach(i -> {
            List<MyRow> myRows = new ArrayList<>();
            initExcelData(i, myRows);
            // 模型名称
            String modelName = toUnderlineCase(i.getSimpleName());
            createSheet(workbook, myRows, modelName);
        });
        return workbook;
    }

    public XSSFSheet createSheet(XSSFWorkbook workbook, List<MyRow> myRows, String modelName) {
        XSSFSheet sheet = workbook.createSheet("sheet" + SHEET_LEVEL++);

        // 创建表头
        XSSFRow modelHeadRow = sheet.createRow(0);
        modelHeadRow.createCell(0).setCellValue("表名");
        modelHeadRow.createCell(1).setCellValue("表描述");
        XSSFRow modelHeadValueRow = sheet.createRow(1);
        modelHeadValueRow.createCell(0).setCellValue(modelName);

        // 创建字段信息行
        XSSFRow modelItemRow = sheet.createRow(2);
        modelItemRow.createCell(0).setCellValue("字段名");
        modelItemRow.createCell(1).setCellValue("字段描述");
        modelItemRow.createCell(2).setCellValue("是否键值");
        modelItemRow.createCell(3).setCellValue("字段属性");
        modelItemRow.createCell(4).setCellValue("备注");
        modelItemRow.createCell(5).setCellValue("枚举值");
        modelItemRow.createCell(6).setCellValue("对象");
        modelItemRow.createCell(7).setCellValue("同步");

        // 遍历要录入的数据
        for (int rowNumber = 3, index = 0; index < myRows.size(); index++, rowNumber++) {
            MyRow myRow = myRows.get(index);
            XSSFRow row = sheet.createRow(rowNumber);
            row.createCell(0).setCellValue(myRow.getFieldName());
            row.createCell(1).setCellValue(myRow.getFieldDesc());
            row.createCell(2).setCellValue("N");
            row.createCell(3).setCellValue(myRow.getFieldType());
            row.createCell(4).setCellValue(myRow.getRemark());
            row.createCell(5).setCellValue(myRow.getEnumValues());
            row.createCell(6).setCellValue(myRow.getLinkObj());
            row.createCell(7).setCellValue(myRow.getSyncType());

        }

        System.out.println(" 模型: " + modelName + " 创建成功");

        return sheet;
    }

    /**
     * 将 Excel 数据写入文件
     *
     * @param workBook Excel 工作簿对象
     * @throws IOException 文件操作异常
     */
    public static void writeExcelToFile(Workbook workBook) throws IOException {
        String filename = "out-" + LocalDateTime.now() + new Random().nextInt() +".xlsx";
        // 指定创建的 Excel 文件名称
        try (BufferedOutputStream outputStream = new BufferedOutputStream(new FileOutputStream(filename))) {
            // 写入数据
            workBook.write(outputStream);
        }
    }

    public static void main(String[] args) throws Exception {
        // 包名前2层即可
        FieldExcelExporter.packageNamePrefix = "com.example";
        FieldExcelExporter.projectName = "TB2B";

        // class列表
        List<Class<?>> classes = List.of(
                MyEntity.class
        );

        for (val c : classes) {
            FieldExcelExporter exporter = FieldExcelExporter.builder()
                    .modelName("_init_")
                    .myRows(new ArrayList<>())
                    .build();

            // 设置实体类
            exporter.initExcelData(c, exporter.myRows);
            Workbook workBook = exporter.createWorkbook();
            writeExcelToFile(workBook);
        }

    }
}
