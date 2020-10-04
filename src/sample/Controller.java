package sample;

import javafx.fxml.FXML;
import javafx.scene.control.TextField;
import javafx.stage.FileChooser;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.net.MalformedURLException;
import java.net.URL;
import java.nio.channels.FileChannel;
import java.nio.file.FileAlreadyExistsException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class Controller {

    @FXML
    private TextField tf_choose_file;

    private static List<String> fileList = new ArrayList<>();


    private static void convertToZip(File dir) {
        boolean isExistArchive = false;
        try {
            File dirSrc = new File("src/");
            File[] files = dirSrc.listFiles();
            if (files != null && files.length > 0) {
                for (File file : files) {
                    if (file.getName().equals("candidats.zip")) {
                        isExistArchive = true;
                        break;
                    }
                }
            }
            if (!isExistArchive) {
                System.out.println("\nStart compressing...");
                compressDirectory(dir.toString());
                System.out.println("\nFinish compressing...");
            }
        } catch (Exception ex) {
            throw new RuntimeException(ex.getMessage(), ex);
        }
    }

    private static void getFileList(File directory) {
        File[] files = directory.listFiles();
        if (files != null && files.length > 0) {
            for (File file : files) {
                if (file.isFile()) {
                    fileList.add(file.getAbsolutePath());
                } else {
                    getFileList(file);
                }
            }
        }
    }

    private static void compressDirectory(String dir) throws IOException {
        File directory = new File(dir);
        getFileList(directory);

        try (FileOutputStream fos = new FileOutputStream("src/candidats.zip");
             ZipOutputStream zos = new ZipOutputStream(fos)) {

            for (String filePath : fileList) {
                System.out.println("Compressing: " + filePath);

                String compressedNameOfDirectory = filePath.substring(
                        directory.getAbsolutePath().length() + 1);

                ZipEntry zipEntry = new ZipEntry(compressedNameOfDirectory);
                zos.putNextEntry(zipEntry);

                try (FileInputStream fis = new FileInputStream(filePath)) {
                    byte[] buffer = new byte[1024];
                    int length;
                    while ((length = fis.read(buffer)) > 0) {
                        zos.write(buffer, 0, length);
                    }
                    zos.closeEntry();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     *
     * @param url Адресс открытой ссылки для скачивания фото с Google Disk
     * @param lastName Фамилия кандидата
     * @param dirNameByTelephone Телефон кандидата в формате без первых символов (+7 или 8)
     * @return Уникальное название файла, хранящее фото кандидата
     * @throws IOException Ошибка при открытии/созданиия/записи изображения, файла или директории
     */
    private static String loadImageFromUrl(URL url, String lastName, String dirNameByTelephone) throws IOException {
        String fileId = url.getQuery().substring(url.getQuery().indexOf("=") + 1);//google file id to download
        String requestDownloadImageUrl = String.format("https://drive.google.com/uc?export=download&id=%s", fileId);
        URL url1 = new URL(requestDownloadImageUrl);
        Path createdPath = Files.createDirectory(Paths.get(String.format("src/candidats/%s", dirNameByTelephone)));
        String fileName = String.format("%s/photo%s.jpg", createdPath.toString(), lastName);
        System.out.println("fileName = " + fileName);
        System.out.println("fileId = " + fileId);
        System.out.println("url for download Image = " + requestDownloadImageUrl);
        System.out.println("----------------------------------------------------");

        try {
            BufferedImage image;
            image = ImageIO.read(url1);
            if (image != null) {
                ImageIO.write(image, "jpg", new File(fileName));
                image.flush();
                return fileName;
            }
        } catch (FileNotFoundException ex) {
            throw new RuntimeException("Error saving image with image name = '" + fileName + "'", ex);
        }
        return null;
    }

    private static void copyFileUsingChannel(File source, File dest) throws IOException {
        try (FileChannel sourceChannel = new FileInputStream(source).getChannel(); FileChannel destChannel = new FileOutputStream(dest).getChannel()) {
            destChannel.transferFrom(sourceChannel, 0, sourceChannel.size());
        }
    }

    @FXML
    public void convert_click() {
        File file = new File("src/candidats/ancets.xlsx");
        System.out.println("candidats.xlsx copy to candidatsCopy.xlxs");
        File fileCopy = new File("src/candidats/ancetsCopy.xlsx");
        try {
            copyFileUsingChannel(file, fileCopy);
        } catch (IOException e) {
            throw new RuntimeException(e.getMessage(), e);
        }
        FileInputStream fip;
        FileInputStream fipCopy;
        try {
            fip = new FileInputStream(file);
            fipCopy = new FileInputStream(fileCopy);
            fileCopy.setWritable(true);
        } catch (FileNotFoundException ex) {
            throw new RuntimeException(ex.getMessage(), ex);
        }

        XSSFWorkbook workbook;
        XSSFWorkbook workbookCopy;
        try {
            workbook = new XSSFWorkbook(fip);
            workbookCopy = new XSSFWorkbook(fipCopy);
        } catch (IOException ex) {
            throw new RuntimeException(ex.getMessage(), ex);
        }
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFSheet sheetCopy = workbookCopy.getSheetAt(0);
        try {
            try {
                Files.createDirectory(Paths.get("src/candidats/files"));
            } catch (FileAlreadyExistsException ex) {
                System.out.println("Directory /files is already exist " + ex.getMessage());
            } catch (IOException e) {
                throw new RuntimeException(e.getMessage(), e);
            }

            if (file.isFile() && file.exists()) {
                System.out.println("candidats.xlsx open");
                System.out.println("candidatsCopy.xlsx open");
                Iterator<Row> iter = sheet.iterator();
                Iterator<Row> iterCopy = sheetCopy.iterator();
                iter.next();
                iterCopy.next();

                while (iter.hasNext() && iterCopy.hasNext()) {
                    Row currentRow = iter.next();
                    Row currentRowCopy = iterCopy.next();
                    if (currentRow.getCell(2) == null || currentRowCopy.getCell(2) == null) {//пустая строка
                        continue;
                    }
                    String urlPhoto = currentRow.getCell(2).getStringCellValue();
                    String lastName = currentRow.getCell(3).getStringCellValue();
                    String telephone = currentRow.getCell(6).getStringCellValue().substring(2);
                    assert urlPhoto != null;
                    assert lastName != null;

                    URL url;
                    try {
                        url = new URL(urlPhoto);
                    } catch (MalformedURLException ex) {
                        throw new RuntimeException(ex.getMessage(), ex);
                    }

                    try {
                        String photoFullPath = loadImageFromUrl(url, lastName, "files/" + telephone);
                        System.out.println("До = " + currentRowCopy.getCell(2));
                        currentRowCopy.getCell(2).setCellValue(photoFullPath);
                        System.out.println("После  = " + currentRowCopy.getCell(2).getStringCellValue());
                    } catch (FileAlreadyExistsException fileEx) {
                        System.out.println("Directory is already exist " + fileEx.getMessage());
                        continue;
                    } catch (IOException ex) {
                        throw new RuntimeException(ex.getMessage(), ex);
                    }
                }

                FileOutputStream f1 = new FileOutputStream("src/candidats/ancetsCopy.xlsx");
                workbookCopy.write(f1);
                f1.close();
                fipCopy.close();
                fip.close();

                File file1 = new File("src/candidats");
                convertToZip(file1);

            } else {
                System.out.println("candidats.xlsx either not exist or can't open");
            }
        } catch (Exception ex) {
            throw new RuntimeException("Global Exception" + ex.getMessage(), ex);
        }
    }

    @FXML
    public void choose_click() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Open Resource File");
        FileChooser.ExtensionFilter xlsxFilter = new FileChooser.ExtensionFilter("XLSX files (*.xlsx)", "*.xlsx");
        FileChooser.ExtensionFilter xlsFilter = new FileChooser.ExtensionFilter("XLS files (*.xls)", "*.xls");
        fileChooser.getExtensionFilters().add(xlsxFilter);
        fileChooser.getExtensionFilters().add(xlsFilter);
        File file = fileChooser.showOpenDialog(null);
        System.out.println(file);
        tf_choose_file.setText(file.getAbsolutePath());

        boolean isExistExcelFile = false;
        boolean isExistExcelFileCopy = false;
        try {
            File dirSrc = new File("src/candidats/");
            File[] files = dirSrc.listFiles();
            if (files != null && files.length > 0) {
                for (File file1 : files) {
                    if (file1.getName().equals("ancets.xlsx")) {
                        isExistExcelFile = true;
                    }
                    if (file1.getName().equals("ancetsCopy.xlsx")) {
                        isExistExcelFileCopy = true;
                    }
                }
            } else {
                try {
                    Files.createDirectory(Paths.get("src/candidats/"));
                } catch (FileAlreadyExistsException ex) {
                    System.out.println("Directory candidats already exist");
                }
            }
            if (isExistExcelFile) {
                System.out.println("\nStart delete ancets.xlsx");
                File fileExcelToDelete = new File("src/candidats/ancets.xlsx");
                boolean isDeleted = fileExcelToDelete.delete();
                System.out.println("Finish delete... = " + isDeleted);
            }
            if (isExistExcelFileCopy) {
                System.out.println("\nStart delete ancetsCopy.xlsx.");
                File fileExcelCopyToDelete = new File("src/candidats/ancetsCopy.xlsx");
                boolean isDeleted = fileExcelCopyToDelete.delete();
                System.out.println("Finish delete ... = " + isDeleted);
            }
        } catch (Exception ex) {
            throw new RuntimeException(ex.getMessage(), ex);
        }

        File copy = new File("src/candidats/" + "ancets.xlsx");
        try (FileInputStream fis = new FileInputStream(file)) {
            FileOutputStream fos = new FileOutputStream(copy);
            byte[] buffer = new byte[1024];
            int length;
            while ((length = fis.read(buffer)) > 0) {
                fos.write(buffer, 0, length);
            }
            fos.close();
        } catch (IOException ex) {
            throw new RuntimeException(ex.getMessage(), ex);
        }
        System.out.println(copy);
    }
}
