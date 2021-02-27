
import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import org.apache.poi.POITextExtractor;
import org.apache.poi.extractor.ExtractorFactory;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.impl.piccolo.io.FileFormatException;
import service.MainFrame;

import javax.swing.*;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.*;
import java.math.RoundingMode;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.util.*;
import java.util.stream.Collectors;

public class GoldMaker {
    static JFrame jFrame;

    public static void main(String args[]) {
        List<Double> valuesGot = new ArrayList<>();
        jFrame = new JFrame();
        jFrame.setExtendedState(JFrame.MAXIMIZED_BOTH);
        jFrame.setVisible(true);

        String replacedFileName="C:\\Users\\768901\\Downloads\\fff.docx";
        JPanel p1 = new JPanel();

        p1.setLayout(new BoxLayout(p1, BoxLayout.X_AXIS));
        JPanel p2 = new JPanel();
        JPanel p3 = new JPanel();
        JPanel p4 = new JPanel();
        JTabbedPane tp = new JTabbedPane();


        JList<File> displayList = new JList<>(new File("C:\\Users\\768901\\Downloads").listFiles());
        displayList.setSelectionMode(javax.swing.ListSelectionModel.SINGLE_SELECTION);
        displayList.setCellRenderer(new MainFrame());
        displayList.setName("displayList");
        displayList.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                String s = displayList.getSelectedValue().toString();
                s.replaceAll("\\\\", "\\\\\\\\");
                System.out.println(displayList.getSelectedValue());
                if (s.endsWith("docx") || s.endsWith("doc")) {
                    FileInputStream in = null;
                    try {
                        Document doc = new Document(s);

                        doc.save(replacedFileName, SaveFormat.DOCX);
                    } catch (Exception ex) {
                        ex.printStackTrace();
                    }
                    try {
                        FileInputStream file = new FileInputStream(replacedFileName);
                        POITextExtractor extractor = ExtractorFactory.createExtractor(file);

                        XWPFDocument doc = new XWPFDocument(
                                OPCPackage.open(replacedFileName));

                        System.out.println(doc.getParagraphs().size()+ "SIZE");

                        for (XWPFParagraph p : doc.getParagraphs()) {
                            List<XWPFRun> runs = p.getRuns();
                            if (runs != null) {
                                for (XWPFRun r : runs) {
                                    String text = r.getText(0);
                                    if (text != null && text.contains("Gold")) {
                                        System.out.println("TOKEN");
                                        StringTokenizer st = new StringTokenizer(text);
                                        List<String> str = new ArrayList<>();
                                        while (st.hasMoreTokens()) {
                                            str.add(st.nextToken());
                                        }
                                        System.out.println("fjoejfoejfoeoejfjeofejofjejfoejfoejofeojfoeofe" + str.get(3));
                                        Double l = Double.parseDouble(str.get(3).trim());
                                        valuesGot.add(l);
                                    }
                                    if (text != null && text.contains("Evaluation")) {
                                        text = text.replace(text, "");
                                    }
                                }
                            }
                        }


                        System.out.println(extractor.getText());
                    } catch (Exception ex) {
                        ex.printStackTrace();
                    }

                    try {
                        String fileName = replacedFileName;

                        XWPFDocument doc = new XWPFDocument(new FileInputStream(fileName));

                        List<XWPFTable> table = doc.getTables();
                        DecimalFormat df = new DecimalFormat("#.###");
                        df.setRoundingMode(RoundingMode.CEILING);

                        for (XWPFTable xwpfTable : table) {
                            System.out.println("Silver" + xwpfTable.getRow(0).getCell(1).getText());
                            System.out.println("Copper" + xwpfTable.getRow(0).getCell(3).getText());
                            System.out.println("ZINC" + xwpfTable.getRow(0).getCell(5).getText());
                            System.out.println("CADIMIUM" + xwpfTable.getRow(1).getCell(1).getText());
                            System.out.println("NICKEL" + xwpfTable.getRow(1).getCell(3).getText());
                            System.out.println("PALLADIUM" + xwpfTable.getRow(1).getCell(5).getText());
                            System.out.println("INDIUM" + xwpfTable.getRow(2).getCell(1).getText());
                            System.out.println("TIN" + xwpfTable.getRow(2).getCell(3).getText());
                            System.out.println("IREDIUM" + xwpfTable.getRow(2).getCell(5).getText());
                            System.out.println("RUTHENIUM" + xwpfTable.getRow(3).getCell(1).getText());
                            System.out.println("TUNGESTAN" + xwpfTable.getRow(3).getCell(3).getText());
                            System.out.println("PLATINUM" + xwpfTable.getRow(3).getCell(5).getText());

                            for(int i=0;i<=6;i++){
                                for(int j=1;j<=5;j++){
                                    try {
                                        valuesGot.add(Double.parseDouble(xwpfTable.getRow(i).getCell(j).getText().trim()));
                                    } catch (NumberFormatException ex) {
                                        ex.printStackTrace();
                                    }
                                    j++;
                                }
                            }

                            System.out.println(valuesGot);
                            System.out.println("LEAD" + xwpfTable.getRow(4).getCell(1).getText());
                            System.out.println("CROMIUM" + xwpfTable.getRow(4).getCell(3).getText());
                            System.out.println("MANGANESE" + xwpfTable.getRow(4).getCell(5).getText());
                            System.out.println("RHODIUM" + xwpfTable.getRow(5).getCell(1).getText());
                            System.out.println("IRON" + xwpfTable.getRow(5).getCell(3).getText());
                            System.out.println("COBALT" + xwpfTable.getRow(5).getCell(5).getText());
                            System.out.println("GALLIUM" + xwpfTable.getRow(6).getCell(1).getText());
                            System.out.println("Vanadium" + xwpfTable.getRow(6).getCell(3).getText());
                            System.out.println("OSMIUM" + xwpfTable.getRow(6).getCell(5).getText());

                        }
                        double sum=valuesGot.stream().mapToDouble(f -> f.doubleValue()).sum();
                        System.out.println("sum "+sum);

                        if(sum<100){

                            double totalLess=100-sum;
                            System.out.println("totalLess "+ totalLess);
                            double silver = 0;
                            try {
                                 silver = Double.parseDouble(table.get(0).getRow(0).getCell(1).getText().trim());
                            }catch(Exception el){
                                el.printStackTrace();
                            }
                            System.out.println("//////////////////////////////////////////");
                            if(silver<=0){
                                System.out.println("SIlverr "+silver);
                                try(FileOutputStream out = new FileOutputStream(new File(replacedFileName))) {
                                    table.get(0).getRow(0).getCell(1).removeParagraph(0);
                                    table.get(0).getRow(0).getCell(1).setText(String.valueOf(df.format(totalLess)));
                                    doc.write(out);
                                }
                                catch(Exception er){
                                    er.printStackTrace();
                                }
                            }else{
                                try(FileOutputStream out = new FileOutputStream(new File(replacedFileName))) {
                                    double dd=silver+totalLess;
                                    System.out.println("siver +less "+ dd);

                                    table.get(0).getRow(0).getCell(1).removeParagraph(0);
                                    table.get(0).getRow(0).getCell(1).setText(String.valueOf(df.format(dd)));
                                    doc.write(out);
                                    System.out.println(table.get(0).getRow(0).getCell(1).getText());
                                }catch(Exception er){
                                    er.printStackTrace();
                                }
                            }
                        }
                        else if(sum>100){

                            double totalLess=sum-100;
                            System.out.println("totalLess "+ totalLess);
                            double silver = 0;
                            try {
                                silver = Double.parseDouble(table.get(0).getRow(0).getCell(1).getText().trim());
                            }catch(Exception el){
                                el.printStackTrace();
                            }
                            System.out.println("//////////////////////////////////////////");
                            if(silver<=0){
                                System.out.println("SIlverr "+silver);
                                try(FileOutputStream out = new FileOutputStream(new File(replacedFileName))) {
                                    table.get(0).getRow(0).getCell(1).removeParagraph(0);
                                    table.get(0).getRow(0).getCell(1).setText(String.valueOf(df.format(totalLess)));
                                    doc.write(out);
                                }
                                catch(Exception er){
                                    er.printStackTrace();
                                }
                            }else{
                                try(FileOutputStream out = new FileOutputStream(new File(replacedFileName))) {
                                    double dd=silver-totalLess;
                                    System.out.println("siver +less "+ dd);

                                    table.get(0).getRow(0).getCell(1).removeParagraph(0);
                                    table.get(0).getRow(0).getCell(1).setText(String.valueOf(df.format(dd)));
                                    doc.write(out);
                                    System.out.println(table.get(0).getRow(0).getCell(1).getText());
                                }catch(Exception er){
                                    er.printStackTrace();
                                }
                            }
                        }

                    } catch (FileFormatException et) {
                        et.printStackTrace();
                    } catch (FileNotFoundException ee) {
                        ee.printStackTrace();
                    } catch (IOException ef) {
                        ef.printStackTrace();
                    }


                } else {
                    JOptionPane.showMessageDialog(jFrame, "Wrong FIle");
                }


            }
        });
        JScrollPane sc = new JScrollPane(displayList);

        sc.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);
        sc.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
        p1.add(sc);
        p1.add(p2);
        p1.add(p3);
        p1.add(p4);

        tp.add("Pure Gold", p1);
        jFrame.add(tp);
    }

    public static List<String> readFileIntoListOfWords(String f) {
        try {
            return Files.readAllLines(Paths.get(f))
                    .stream()
                    .map(l -> l.split(" "))
                    .flatMap(Arrays::stream)
                    .collect(Collectors.toList());
        } catch (IOException e) {
            e.printStackTrace();
        }
        return Collections.emptyList();
    }
    private static void removeParagraphs(XWPFTableCell tableCell) {
        int count = tableCell.getParagraphs().size();
        for(int i = 0; i < count; i++){
            tableCell.removeParagraph(i);
        }
    }

}
