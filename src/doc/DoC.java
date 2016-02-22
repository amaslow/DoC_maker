package doc;

import java.awt.Desktop;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileFilter;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.imageio.ImageIO;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.awt.Color;
import java.text.ParseException;
import java.util.ArrayList;

public class DoC extends javax.swing.JFrame {

    File productContent = new File("G:\\Product Content\\PRODUCTS");
    File images = new File("G:\\QC\\Database\\Items\\Reports\\");
    File docFolder = new File("H:\\DoCs_temporary\\");
    Connection con = null;
    Statement st = null;
    ResultSet rs = null;
    String sap, item, brand, descr_en, descr_fr, descr_de;
    String emc_ce, emc1, emc2, emc3, emc4, emc5, emc6, emc7, emc8, emc9, emc10;
    String lvd_ce, lvd1, lvd2, lvd3, lvd4, lvd5, lvd6, lvd7, lvd8, lvd9;
    String rf_ce, rf1, rf2, rf3, rf4, rf_f;
    String eup_ce;
    String cpd_dir, cpd1, cpd2, cpd3, cpd4;
    String rohs_ce;
    String component1, component2, component3, component4, component5, component6, component7, component8, component9, component10;
    Desktop desktop = null;

    public DoC() {
        initComponents();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        sapField = new javax.swing.JFormattedTextField();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        dateField = new com.toedter.calendar.JDateChooser();
        jButton1 = new javax.swing.JButton();
        workLabel = new javax.swing.JLabel();
        signComboBox = new javax.swing.JComboBox();
        signLabel = new javax.swing.JLabel();
        jButton2 = new javax.swing.JButton();
        jButton3 = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("DoC maker");

        try {
            sapField.setFormatterFactory(new javax.swing.text.DefaultFormatterFactory(new javax.swing.text.MaskFormatter("##.###.##")));
        } catch (java.text.ParseException ex) {
            ex.printStackTrace();
        }
        sapField.setHorizontalAlignment(javax.swing.JTextField.CENTER);
        sapField.setToolTipText("type SAP number here");
        sapField.setFont(new java.awt.Font("Tahoma", 0, 24)); // NOI18N

        jLabel1.setText("SAP:");

        jLabel2.setText("Date:");

        dateField.setDate(new Date());
        dateField.setDateFormatString("dd-MM-yyyy");
        dateField.setFont(new java.awt.Font("Tahoma", 0, 18)); // NOI18N
        dateField.addPropertyChangeListener(new java.beans.PropertyChangeListener() {
            public void propertyChange(java.beans.PropertyChangeEvent evt) {
                dateFieldPropertyChange(evt);
            }
        });

        jButton1.setText("Make DoC");
        jButton1.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mousePressed(java.awt.event.MouseEvent evt) {
                jButton1MousePressed(evt);
            }
        });
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });

        workLabel.setForeground(new java.awt.Color(0, 51, 255));

        signComboBox.setModel(new javax.swing.DefaultComboBoxModel(new String[] { "Jose", "Ad" }));

        signLabel.setText("signature:");

        jButton2.setText("move pdf to Product Content");
        jButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton2ActionPerformed(evt);
            }
        });

        jButton3.setText("create pdf's");
        jButton3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton3ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addGap(0, 0, Short.MAX_VALUE)
                .addComponent(workLabel, javax.swing.GroupLayout.PREFERRED_SIZE, 67, javax.swing.GroupLayout.PREFERRED_SIZE))
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jButton3)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(jButton2))
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(sapField, javax.swing.GroupLayout.DEFAULT_SIZE, 125, Short.MAX_VALUE)
                            .addComponent(dateField, javax.swing.GroupLayout.DEFAULT_SIZE, 125, Short.MAX_VALUE))
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addGap(42, 42, 42)
                                .addComponent(jButton1, javax.swing.GroupLayout.PREFERRED_SIZE, 94, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(signLabel)))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(signComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(0, 66, Short.MAX_VALUE)))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap(30, Short.MAX_VALUE)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(sapField, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(jButton1, javax.swing.GroupLayout.DEFAULT_SIZE, 35, Short.MAX_VALUE)
                            .addComponent(jLabel2, javax.swing.GroupLayout.DEFAULT_SIZE, 35, Short.MAX_VALUE)
                            .addComponent(dateField, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
                    .addGroup(javax.swing.GroupLayout.Alignment.LEADING, layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(signComboBox, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(signLabel)))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jButton2)
                    .addComponent(jButton3))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(workLabel))
        );

        pack();
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        String sapNo = sapField.getText();
        final String sapWithoutDocs = sapNo.replace(".", "");
        Date Date = dateField.getDate();
        File newDocPath = new File(productContent + "\\" + sapWithoutDocs + "\\");
        File newDocFile = null;

        if (!sapNo.equals("  .   .  ")) {
            if (newDocPath.exists()) {
                File[] allDocFiles = newDocPath.listFiles(new FileFilter() {

                    @Override
                    public boolean accept(File pathname) {
                        return pathname.isFile() && pathname.getName().toString().startsWith("DoC_" + sapWithoutDocs);
                    }
                });
                switch (allDocFiles.length) {
                    case 0:
                        newDocFile = null;
                        break;
                    case 1:
                        newDocFile = allDocFiles[0];
                        break;
                }
            } else {
                newDocPath.mkdirs();
            }

            if (Desktop.isDesktopSupported()) {
                desktop = Desktop.getDesktop();
            }
            if (newDocFile != null) {
                String createDay = newDocFile.getName().toString().substring(18, 20);
                String createMonth = newDocFile.getName().toString().substring(16, 18);
                String createYear = newDocFile.getName().toString().substring(12, 16);

                Object[] options = {"Open existing DoC file", "Create a new DoC file"};
                int n = JOptionPane.showOptionDialog(null, "DoC file already exist.\n" + "Created on: " + createDay + "." + createMonth + "." + createYear, "DoC",
                        JOptionPane.DEFAULT_OPTION, JOptionPane.QUESTION_MESSAGE, null, options, options[0]);
                if (n == JOptionPane.YES_OPTION) {
                    try {
                        if (newDocFile.exists()) {
                            desktop.open(newDocFile);
                        }
                    } catch (IOException e) {
                    }
                } else if (n == JOptionPane.NO_OPTION) {
                    try {
                        createDocExcel(Date, sapNo);
                    } catch (IOException ex) {
                        Logger.getLogger(DoC.class.getName()).log(Level.SEVERE, null, ex);
                    } finally {
                        Utils.closeDB(rs, st, con);
                    }
                }
            } else {
                Object[] optionsDoC = {"Create a DoC in PDF and save in Product Content", "Create and open a DoC in Excel"};
                int a = JOptionPane.showOptionDialog(null, "DoC file already exist.", "DoC",
                        JOptionPane.DEFAULT_OPTION, JOptionPane.QUESTION_MESSAGE, null, optionsDoC, optionsDoC[0]);
                if (a == JOptionPane.YES_OPTION) {
                    try {
                        createDocExcel(Date, sapNo);
                    } catch (IOException ex) {
                        Logger.getLogger(DoC.class.getName()).log(Level.SEVERE, null, ex);
                    } finally {
                        Utils.closeDB(rs, st, con);
                    }
                }
            }
        } else {
            JOptionPane.showMessageDialog(null, "Enter SAP number !!!", "No SAP", JOptionPane.ERROR_MESSAGE);
            sapField.requestFocus();
        }

    }//GEN-LAST:event_jButton1ActionPerformed

    private void jButton1MousePressed(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jButton1MousePressed
        workLabel.setText("working...");
        workLabel.setForeground(Color.red);
    }//GEN-LAST:event_jButton1MousePressed

    private void dateFieldPropertyChange(java.beans.PropertyChangeEvent evt) {//GEN-FIRST:event_dateFieldPropertyChange

        Date date = dateField.getDate();
        Date today = new Date();
        date.setTime(date.getTime() + 5000);

        if (date.before(today)) {
            signComboBox.setSelectedIndex(1);
        } else {
            signComboBox.setSelectedIndex(0);
        }
    }//GEN-LAST:event_dateFieldPropertyChange

    private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton2ActionPerformed
        File productContent = new File("G:\\Product Content\\PRODUCTS\\");
        File[] oldfiles = docFolder.listFiles(new FileFilter() {
            @Override
            public boolean accept(File pathname) {
                return pathname.isFile() && pathname.getName().toString().endsWith(".pdf");
            }
        });

        if (oldfiles.length > 0) {

            for (int i = 0; i < oldfiles.length; i += 1) {
                String sap = oldfiles[i].getName().substring(4, 11);
                File newFile = new File(productContent + "\\" + sap + "\\" + oldfiles[i].getName());
                System.out.println(oldfiles[i] + " - " + sap + " - " + newFile);
                oldfiles[i].renameTo(newFile);
            }
            JOptionPane.showMessageDialog(null, "All DoCs have been moved into proper Product Content folders", "Done", JOptionPane.PLAIN_MESSAGE);
        } else {
            JOptionPane.showMessageDialog(null, "There is no DoC in pdf file", "No pdf DoC", JOptionPane.ERROR_MESSAGE);
        }
    }//GEN-LAST:event_jButton2ActionPerformed

    private void jButton3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton3ActionPerformed
        createPDF();
    }//GEN-LAST:event_jButton3ActionPerformed

    public static void main(String args[]) {

        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new DoC().setVisible(true);
            }
        });
    }
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private com.toedter.calendar.JDateChooser dateField;
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton3;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JFormattedTextField sapField;
    private javax.swing.JComboBox signComboBox;
    private javax.swing.JLabel signLabel;
    private javax.swing.JLabel workLabel;
    // End of variables declaration//GEN-END:variables

    public ArrayList<String> isSet(String sapNo) {
        ArrayList compList = new ArrayList<String>();
        try {
            String sql_SET = "SELECT "
                    + "SAP,"
                    + "COALESCE(COMPONENT1,''),COALESCE(COMPONENT2,''),COALESCE(COMPONENT3,''),COALESCE(COMPONENT4,''),COALESCE(COMPONENT5,''),COALESCE(COMPONENT6,''),COALESCE(COMPONENT7,''),COALESCE(COMPONENT8,''),COALESCE(COMPONENT9,''),COALESCE(COMPONENT10,'')"
                    + " FROM elro.items"
                    + " WHERE SAP='" + sapNo + "';";

            con = Utils.getConnection();

            st = con.createStatement();
            rs = st.executeQuery(sql_SET);

            while (rs.next()) {
                sap = rs.getString(1);
                component1 = rs.getString(2);
                if (component1.length() > 0) {
                    compList.add(component1.substring(0, 9));
                }
                component2 = rs.getString(3);
                if (component2.length() > 0) {
                    compList.add(component2.substring(0, 9));
                }
                component3 = rs.getString(4);
                if (component3.length() > 0) {
                    compList.add(component3.substring(0, 9));
                }
                component4 = rs.getString(5);
                if (component4.length() > 0) {
                    compList.add(component4.substring(0, 9));
                }
                component5 = rs.getString(6);
                if (component5.length() > 0) {
                    compList.add(component5.substring(0, 9));
                }
                component6 = rs.getString(7);
                if (component6.length() > 0) {
                    compList.add(component6.substring(0, 9));
                }
                component7 = rs.getString(8);
                if (component7.length() > 0) {
                    compList.add(component7.substring(0, 9));
                }
                component8 = rs.getString(9);
                if (component8.length() > 0) {
                    compList.add(component8.substring(0, 9));
                }
                component9 = rs.getString(10);
                if (component9.length() > 0) {
                    compList.add(component9.substring(0, 9));
                }
                component10 = rs.getString(11);
                if (component10.length() > 0) {
                    compList.add(component10.substring(0, 9));
                }
            }

        } catch (Exception X) {
        } finally {
            Utils.closeDB(rs, st, con);
        }
        return compList;
    }

    private void createDocExcel(Date Date, String sapNo) throws FileNotFoundException, IOException {
        workLabel.setText("working...");
        String sapWithoutDocs = sapNo.replace(".", "");
        DateFormat DateSapFormat = new SimpleDateFormat("yyyyMMdd");
        if (!docFolder.exists()) {
            docFolder.mkdirs();
        }

        File excelname = new File(docFolder + "\\DoC_" + sapWithoutDocs + "_" + DateSapFormat.format(Date).toString() + ".xlsx");
        FileOutputStream fos = new FileOutputStream(excelname);
        XSSFWorkbook wb = new XSSFWorkbook();

        if (isSet(sapNo).size() < 1) {
            wb.createSheet("EN");
            wb.createSheet("DE");
            wb.createSheet("FR");
            int count = 0;
            createDoC(Date, sapNo, count, wb);
        } else {
            wb.createSheet("Intro");
            createIntro(Date, sapNo, wb);
            for (int i = 0; i < isSet(sapNo).size(); i += 1) {
                wb.createSheet("EN" + (i + 1));
                wb.createSheet("DE" + (i + 1));
                wb.createSheet("FR" + (i + 1));
                int count = (isSet(sapNo).indexOf(isSet(sapNo).get(i)) * 3) + 1;
                createDoC(Date, isSet(sapNo).get(i).toString(), count, wb);
            }
        }

        wb.write(fos);
        fos.flush();
        fos.close();

        workLabel.setText("done");
        workLabel.setForeground(Color.blue);

        if (Desktop.isDesktopSupported()) {
            desktop = Desktop.getDesktop();
        }
        desktop.open(excelname);

    }

    private void createDoC(Date Date, String sapNo, int count, XSSFWorkbook wb) {
        System.out.println(sapNo);
        DateFormat DateFormat = new SimpleDateFormat("dd-MM-yyyy");
        String sapWithoutDocs = sapNo.replace(".", "");

        CellStyle header = wb.createCellStyle();
        XSSFFont arialBlack = wb.createFont();
        arialBlack.setFontName("Arial Black");
        arialBlack.setFontHeightInPoints((short) 12);
        arialBlack.setBold(true);
        header.setFont(arialBlack);
        header.setAlignment(XSSFCellStyle.ALIGN_CENTER);

        CellStyle itemStyle = wb.createCellStyle();
        itemStyle.setFont(arialBlack);
        itemStyle.setAlignment(XSSFCellStyle.ALIGN_LEFT);

        CellStyle normal = wb.createCellStyle();
        XSSFFont arial = wb.createFont();
        arial.setFontName("Arial");
        arial.setFontHeightInPoints((short) 10);
        normal.setFont(arial);

        CellStyle normalW = wb.createCellStyle();
        normalW.setFont(arial);
        normalW.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        normalW.setWrapText(true);

        CellStyle normalH = wb.createCellStyle();
        XSSFFont arialH = wb.createFont();
        arialH.setFontName("Arial");
        arialH.setFontHeightInPoints((short) 12);
        normalH.setFont(arialH);

        CellStyle normalHI = wb.createCellStyle();
        XSSFFont arialHI = wb.createFont();
        arialHI.setFontName("Arial");
        arialHI.setItalic(true);
        arialHI.setFontHeightInPoints((short) 12);
        normalHI.setFont(arialHI);

        CellStyle normalB = wb.createCellStyle();
        XSSFFont arialB = wb.createFont();
        arialB.setFontName("Arial Black");
        arialB.setFontHeightInPoints((short) 10);
        arialB.setBold(true);
        normalB.setFont(arialB);

        CellStyle normalHW = wb.createCellStyle();
        arialH.setBold(true);
        normalHW.setFont(arialH);
        normalHW.setWrapText(true);
        normalHW.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        normalHW.setAlignment(XSSFCellStyle.ALIGN_CENTER);

        String sql = "SELECT "
                + "SAP,ITEM,BRAND,DESCR_EN,DESCR_FR,DESCR_DE,"
                + "COALESCE(EMC_CE,''),COALESCE(EMC1,''),COALESCE(EMC2,''),COALESCE(EMC3,''),COALESCE(EMC4,''),COALESCE(EMC5,''),COALESCE(EMC6,''),COALESCE(EMC7,''),COALESCE(EMC8,''),COALESCE(EMC9,''),COALESCE(EMC10,''),"
                + "COALESCE(LVD_CE,''),COALESCE(LVD1,''),COALESCE(LVD2,''),COALESCE(LVD3,''),COALESCE(LVD4,''),COALESCE(LVD5,''),COALESCE(LVD6,''),COALESCE(LVD7,''),COALESCE(LVD8,''),COALESCE(LVD9,''),"
                + "COALESCE(RF_CE,''),COALESCE(RF1,''),COALESCE(RF2,''),COALESCE(RF3,''),COALESCE(RF4,''),COALESCE(RF_F,''),"
                + "COALESCE(EUP_CE,''),"
                + "COALESCE(CPD_DIR,''),COALESCE(CPD1,''),COALESCE(CPD2,''),COALESCE(CPD3,''),COALESCE(CPD4,''),"
                + "COALESCE(ROHS_CE,'')"
                + " FROM elro.items"
                + " WHERE SAP='" + sapNo + "';";
        con = Utils.getConnection();

        try {
            st = con.createStatement();
            rs = st.executeQuery(sql);

            while (rs.next()) {
                sap = rs.getString(1);
                item = rs.getString(2);
                brand = rs.getString(3);
                descr_en = rs.getString(4);
                descr_fr = rs.getString(5);
                descr_de = rs.getString(6);
                emc_ce = rs.getString(7);
                emc1 = rs.getString(8);
                emc2 = rs.getString(9);
                emc3 = rs.getString(10);
                emc4 = rs.getString(11);
                emc5 = rs.getString(12);
                emc6 = rs.getString(13);
                emc7 = rs.getString(14);
                emc8 = rs.getString(15);
                emc9 = rs.getString(16);
                emc10 = rs.getString(17);
                lvd_ce = rs.getString(18);
                lvd1 = rs.getString(19);
                lvd2 = rs.getString(20);
                lvd3 = rs.getString(21);
                lvd4 = rs.getString(22);
                lvd5 = rs.getString(23);
                lvd6 = rs.getString(24);
                lvd7 = rs.getString(25);
                lvd8 = rs.getString(26);
                lvd9 = rs.getString(27);
                rf_ce = rs.getString(28);
                rf1 = rs.getString(29);
                rf2 = rs.getString(30);
                rf3 = rs.getString(31);
                rf4 = rs.getString(32);
                rf_f = rs.getString(33);
                eup_ce = rs.getString(34);
                cpd_dir = rs.getString(35);
                cpd1 = rs.getString(36);
                cpd2 = rs.getString(37);
                cpd3 = rs.getString(38);
                cpd4 = rs.getString(39);
                rohs_ce = rs.getString(40);

                String[] declaration = {"EU DECLARATION OF CONFORMITY", "EU-KONFORMITÄTSERKLÄRUNG", "DÉCLARATION UE DE CONFORMITÉ"};
                String[] company = {"Company:", "Firma:", "Société:"};
                String[] address = {"Address, City:", "Addresse:", "Adresse/Ville:"};
                String[] country1 = {"Country:", "Land:", "Pays:"};
                String[] country2 = {"The Netherlands", "Niederlande", "Pays-Bas"};
                String[] resp = {"This declaration of conformity is issued under the sole responsibility of the manufacturer.",
                    "Die alleinige Verantwortung für die Ausstellung dieser Konformitätserklärung trägt der Hersteller.",
                    "La présente déclaration de conformité est établie sous la seule responsabilité du fabricant."};
                String[] object = {"Object of the declaration:", "Gegenstand der Erklärung:", "Objet de la déclaration:"};
                String[] descr = {"Description:", "Artikelbeschreibung:", "Description:"};
                String[] descr_lang = {descr_en, descr_de, descr_fr};
                String[] name = {"Pruduct name:", "Artikel-Nr.:", "Référence produit:"};
                String[] trademark = {"Trademark:", "Markenname:", "Marque déposée:"};
                String[] directives = {"The object of the declaration described above is in conformity with the relevant Union harmonisation legislation:",
                    "Der oben beschriebene Gegenstand der Erklärung erfüllt die einschlägigen Harmonisierungsrechtsvorschriften der Union:",
                    "L’objet de la déclaration décrit ci-dessus est conforme à la législation d’harmonisation de l’Union applicable:"};
                String[] dir_emc = {"Electro Magnetic Compatibility Directive", "Elektromagnetische Verträglichkeit (EMV)", "Compatibilité Electro Magnétique"};
                String[] dir_cpd1 = {"CPD Directive, CPR Regulation (EU)", "CPD Richtlinie, CPR Verordnung", "CPD Directive, CPR Règlement"};
                String[] dir_cpd2 = {"Pressure equipment Directive", "Druckgeräte Richtlinie", "Équipements sous pression Directive"};
                String[] dir_cpd3 = {"General product safety Directive", "Allgemeine Produktsicherheit Richtlinie", "Sécurité générale des produits Directive"};
                String[] dir_cpd4 = {"Personal protective equipment Directive", "Persönliche Schutzausrüstungen Richtlinie", "Équipements Directive"};
                String[] dir_erp1 = {"Ecodesign Directives (ErP)", "Ecodesign Richtlinien (ErP)", "L’écoconception Directives (ErP)"};
                String[] dir_erp2 = {"Ecodesign Regulations (ErP)", "Ecodesign Verordnungen (ErP)", "L’écoconception Règlements (ErP)"};
                String[] dir_lvd = {"Low Voltage Directive", "Niederspannungsrichtlinie (LVD)", "Directive Basse Tension"};
                String[] dir_rf = {"R&TTE Directive", "R&TTE-Anforderung", "Directive R&TTE"};
                String[] dir_rohs = {"RoHS Directive", "Beschränkung der Gefährlicher Stoffe", "Restriction de substances dangereuses"};
                String[] standards = {"References to the relevant harmonised standards used or references to the other technical specifications in relation to which conformity is declared:",
                    "Angabe der einschlägigen harmonisierten Normen, die zugrunde gelegt wurden, oder Angabe der anderen technischen Spezifikationen, in Bezug auf die die Konformität erklärt wird:",
                    "Références des normes harmonisées pertinentes appliquées ou des autres spécifications techniques par rapport auxquelles la conformité est déclarée:"};
                String[] auth_repr = {"Authorized representative:", "Bevollmächtigten Vertreter:", "Représentant autorisé:"};
                String[] place = {"Place and date of issue: Gilze, ", "Ort und Datum der Ausstellung: Gilze, ", "Lieu et date d’établissement: Gilze, "};

                for (int i = 0; i < 3; i += 1) {
                    int rownr = 0;
                    XSSFRow row = wb.getSheetAt(i + count).createRow(rownr);//0
                    XSSFCell cell = row.createCell(0);
                    cell.setCellValue(declaration[i]);
                    cell.setCellStyle(header);
                    wb.getSheetAt(i + count).addMergedRegion(new CellRangeAddress(rownr, rownr, 0, 8));

                    rownr += 2;
                    row = wb.getSheetAt(i + count).createRow(rownr);//2
                    cell = row.createCell(0);
                    cell.setCellValue(company[i]);
                    cell.setCellStyle(normal);
                    cell = row.createCell(2);
                    cell.setCellValue("Smartwares Safety & Lighting B.V.");
                    cell.setCellStyle(normal);

                    FileInputStream logo = new FileInputStream(images + "\\logo.png");
                    byte[] logo_bytes = IOUtils.toByteArray(logo);
                    int logo_id = wb.addPicture(logo_bytes, XSSFWorkbook.PICTURE_TYPE_PNG);
                    logo.close();
                    XSSFDrawing logo_drawing = wb.getSheetAt(i + count).createDrawingPatriarch();
                    XSSFClientAnchor logo_anchor = new XSSFClientAnchor();
                    logo_anchor.setCol1(6);
                    logo_anchor.setRow1(rownr);
                    logo_anchor.setDx1(0);
                    logo_anchor.setDy1(0);
                    logo_anchor.setCol2((short) 8);
                    logo_anchor.setRow2(rownr + 2);
                    logo_anchor.setDx2(0);
                    logo_anchor.setDy2(0);
                    XSSFPicture logo_picture = logo_drawing.createPicture(logo_anchor, logo_id);
                    logo_picture.resize(0.78);

                    rownr += 1;
                    row = wb.getSheetAt(i + count).createRow(rownr);//3
                    cell = row.createCell(0);
                    cell.setCellValue(address[i]);
                    cell.setCellStyle(normal);
                    cell = row.createCell(2);
                    cell.setCellValue("Broekakkerweg 15, 5126 BD Gilze");
                    cell.setCellStyle(normal);

                    rownr += 1;
                    row = wb.getSheetAt(i + count).createRow(rownr);//4
                    cell = row.createCell(0);
                    cell.setCellValue(country1[i]);
                    cell.setCellStyle(normal);
                    cell = row.createCell(2);
                    cell.setCellValue(country2[i]);
                    cell.setCellStyle(normal);

                    rownr += 2;
                    row = wb.getSheetAt(i + count).createRow(rownr);//6
                    cell = row.createCell(0);
                    cell.setCellValue(resp[i]);
                    cell.setCellStyle(normal);

                    rownr += 2;
                    row = wb.getSheetAt(i + count).createRow(rownr);//8
                    cell = row.createCell(0);
                    cell.setCellValue(object[i]);
                    cell.setCellStyle(normal);

                    rownr += 2;
                    row = wb.getSheetAt(i + count).createRow(rownr);//10
                    cell = row.createCell(0);
                    cell.setCellValue(descr[i]);
                    cell.setCellStyle(normal);
                    cell = row.createCell(2);
                    cell.setCellValue(descr_lang[i]);
                    cell.setCellStyle(normalH);

                    rownr += 1;
                    row = wb.getSheetAt(i + count).createRow(rownr);//11
                    cell = row.createCell(0);
                    cell.setCellValue(name[i]);
                    cell.setCellStyle(normal);
                    cell = row.createCell(2);
                    cell.setCellValue(item);
                    cell.setCellStyle(itemStyle);

                    rownr += 1;
                    row = wb.getSheetAt(i + count).createRow(rownr);//12
                    cell = row.createCell(0);
                    cell.setCellValue(trademark[i]);
                    cell.setCellStyle(normal);
                    cell = row.createCell(2);
                    cell.setCellValue(brand.replace("PL ", ""));
                    cell.setCellStyle(normalHI);

                    FileInputStream brand_logo = new FileInputStream(images + "\\BRANDS\\" + brand + ".png");
                    byte[] brand_logo_bytes = IOUtils.toByteArray(brand_logo);
                    int brand_logo_id = wb.addPicture(brand_logo_bytes, XSSFWorkbook.PICTURE_TYPE_PNG);
                    brand_logo.close();
                    XSSFDrawing brand_logo_drawing = wb.getSheetAt(i + count).createDrawingPatriarch();
                    XSSFClientAnchor brand_logo_anchor = new XSSFClientAnchor();
                    brand_logo_anchor.setCol1(5);
                    brand_logo_anchor.setRow1(rownr - 1);
                    XSSFPicture brand_logo_picture = brand_logo_drawing.createPicture(brand_logo_anchor, brand_logo_id);
                    brand_logo_picture.resize();

                    rownr += 2;
                    row = wb.getSheetAt(i + count).createRow(rownr);//14
                    cell = row.createCell(0);
                    cell.setCellValue(directives[i]);
                    cell.setCellStyle(normalW);
                    wb.getSheetAt(i + count).addMergedRegion(new CellRangeAddress(rownr, rownr + 1, 0, 8));

                    rownr += 2;//16

                    if (emc_ce.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(dir_emc[i]);
                        cell.setCellStyle(normal);
                        cell = row.createCell(4);
                        switch (i) {
                            case 1:
                                cell.setCellValue("(" + emc_ce.replace("/EC", "/EG") + ")");
                                break;
                            case 2:
                                cell.setCellValue("(" + emc_ce.replace("/EC", "/CE") + ")");
                                break;
                            default:
                                cell.setCellValue("(" + emc_ce + ")");
                                break;
                        }
                        cell.setCellStyle(normal);
                        rownr += 1;
                    }

                    if (cpd1.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        if (cpd_dir.contains("305/2011")) {
                            cell.setCellValue(dir_cpd1[i]);
                            rownr += 1;
                        } else if (cpd_dir.contains("97/23/EC")) {
                            switch (i) {
                                case 1:
                                    cell.setCellValue(dir_cpd2[i].replace("/EC", "/EG"));
                                    break;
                                case 2:
                                    cell.setCellValue(dir_cpd2[i].replace("/EC", "/CE"));
                                    break;
                                default:
                                    cell.setCellValue(dir_cpd2[i]);
                                    break;
                            }
                            rownr += 1;
                        } else if (cpd_dir.contains("2001/95/EC")) {
                            switch (i) {
                                case 1:
                                    cell.setCellValue(dir_cpd3[i].replace("/EC", "/EG"));
                                    break;
                                case 2:
                                    cell.setCellValue(dir_cpd3[i].replace("/EC", "/CE"));
                                    break;
                                default:
                                    cell.setCellValue(dir_cpd3[i]);
                                    break;
                            }
                            rownr += 1;
                        } else if (cpd_dir.contains("89/686/EEC")) {
                            switch (i) {
                                case 1:
                                    cell.setCellValue(dir_cpd4[i].replace("/EEC", "/EWG"));
                                    break;
                                case 2:
                                    cell.setCellValue(dir_cpd4[i].replace("/EEC", "/CEE"));
                                    break;
                                default:
                                    cell.setCellValue(dir_cpd4[i]);
                                    break;
                            }
                            rownr += 1;
                        }
                        cell.setCellStyle(normal);
                        cell = row.createCell(4);
                        if (cpd_dir.contains("305/2011")) {
                            switch (i) {
                                case 1:
                                    cell.setCellValue("(89/106/EWG, 93/68/EWG, (EU) 305/2011)");
                                    break;
                                case 2:
                                    cell.setCellValue("(89/106/CEE, 93/68/CEE, (EU) 305/2011)");
                                    break;
                                default:
                                    cell.setCellValue("(89/106/EEC, 93/68/EEC, (EU) 305/2011)");
                                    break;
                            }
                        } else if (cpd_dir.contains("97/23/EC")) {
                            switch (i) {
                                case 1:
                                    cell.setCellValue("(97/23/EG)");
                                    break;
                                case 2:
                                    cell.setCellValue("(97/23/CE)");
                                    break;
                                default:
                                    cell.setCellValue("(97/23/EC)");
                                    break;
                            }
                        } else if (cpd_dir.contains("2001/95/EC")) {
                            switch (i) {
                                case 1:
                                    cell.setCellValue("(2001/95/EG)");
                                    break;
                                case 2:
                                    cell.setCellValue("(2001/95/CE)");
                                    break;
                                default:
                                    cell.setCellValue("(2001/95/EC)");
                                    break;
                            }
                        } else if (cpd_dir.contains("89/686/EEC")) {
                            switch (i) {
                                case 1:
                                    cell.setCellValue("(89/686/EWG)");
                                    break;
                                case 2:
                                    cell.setCellValue("(89/686/CEE)");
                                    break;
                                default:
                                    cell.setCellValue("(89/686/EEC)");
                                    break;
                            }
                        }
                        cell.setCellStyle(normal);
                    }

                    if (lvd_ce.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(dir_lvd[i]);
                        cell.setCellStyle(normal);
                        cell = row.createCell(4);
                        switch (i) {
                            case 1:
                                cell.setCellValue("(" + lvd_ce.replace("/EC", "/EG") + ")");
                                break;
                            case 2:
                                cell.setCellValue("(" + lvd_ce.replace("/EC", "/CE") + ")");
                                break;
                            default:
                                cell.setCellValue("(" + lvd_ce + ")");
                                break;
                        }
                        cell.setCellStyle(normal);
                        rownr += 1;
                    }

                    if (eup_ce.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(dir_erp1[i]);
                        cell.setCellStyle(normal);
                        cell = row.createCell(4);
                        if ((eup_ce.contains("1194") && eup_ce.contains("874")) || (eup_ce.contains("244") && eup_ce.contains("874")) || (eup_ce.contains("245") && eup_ce.contains("874"))) {
                            switch (i) {
                                case 1:
                                    cell.setCellValue("(2009/125/EG, 2010/30/EU)");
                                    break;
                                case 2:
                                    cell.setCellValue("(2009/125/CE, 2010/30/EU)");
                                    break;
                                default:
                                    cell.setCellValue("(2009/125/EC, 2010/30/EU)");
                                    break;
                            }
                        } else if (!eup_ce.contains("874")) {
                            switch (i) {
                                case 1:
                                    cell.setCellValue("(2009/125/EG)");
                                    break;
                                case 2:
                                    cell.setCellValue("(2009/125/CE)");
                                    break;
                                default:
                                    cell.setCellValue("(2009/125/EC)");
                                    break;
                            }
                        } else if (eup_ce.contains("874") && !eup_ce.contains("244") && !eup_ce.contains("245") && !eup_ce.contains("1194")) {
                            cell.setCellValue("(2010/30/EU)");
                        }
                        cell.setCellStyle(normal);
                        rownr += 1;
                    }

                    if (eup_ce.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(dir_erp2[i]);
                        cell.setCellStyle(normal);
                        cell = row.createCell(4);
                        cell.setCellValue("(" + eup_ce + ")");
                        cell.setCellStyle(normal);
                        rownr += 1;
                    }

                    if (rf_ce.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(dir_rf[i]);
                        cell.setCellStyle(normal);
                        cell = row.createCell(4);
                        switch (i) {
                            case 1:
                                cell.setCellValue("(" + rf_ce.replace("/EC", "/EG") + ")");
                                break;
                            case 2:
                                cell.setCellValue("(" + rf_ce.replace("/EC", "/CE") + ")");
                                break;
                            default:
                                cell.setCellValue("(" + rf_ce + ")");
                                break;
                        }
                        cell.setCellStyle(normal);
                        rownr += 1;
                    }

                    if (rohs_ce.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(dir_rohs[i]);
                        cell.setCellStyle(normal);
                        cell = row.createCell(4);
                        switch (i) {
                            case 1:
                                cell.setCellValue("(" + rohs_ce.replace("/EEC", "/EWG") + ")");
                                break;
                            case 2:
                                cell.setCellValue("(" + rohs_ce.replace("/EEC", "/CEE") + ")");
                                break;
                            default:
                                cell.setCellValue("(" + rohs_ce + ")");
                                break;
                        }
                        cell.setCellStyle(normal);
                        rownr += 1;
                    }
                    rownr += 1;
                    row = wb.getSheetAt(i + count).createRow(rownr);
                    cell = row.createCell(0);
                    cell.setCellValue(standards[i]);
                    cell.setCellStyle(normalW);
                    wb.getSheetAt(i + count).addMergedRegion(new CellRangeAddress(rownr, rownr + 1, 0, 8));
                    rownr += 2;

                    if (emc1.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(emc1);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (emc2.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(emc2);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (emc3.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(emc3);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (emc4.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(emc4);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (emc5.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(emc5);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (emc6.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(emc6);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (emc7.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(emc7);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (emc8.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(emc8);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (emc9.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(emc9);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (emc10.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(emc10);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (lvd1.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(lvd1);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (lvd2.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(lvd2);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (lvd3.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(lvd3);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (lvd4.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(lvd4);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (lvd5.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(lvd5);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (lvd6.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(lvd6);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (lvd7.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(lvd7);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (lvd8.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(lvd8);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (lvd9.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(lvd9);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (cpd1.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(cpd1);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (cpd2.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(cpd2);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (cpd3.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(cpd3);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (cpd4.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(cpd4);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (rf1.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(rf1);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (rf2.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(rf2);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (rf3.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(rf3);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }
                    if (rf4.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr);
                        cell = row.createCell(0);
                        cell.setCellValue(rf4);
                        cell.setCellStyle(normalB);
                        rownr += 1;
                    }

                    rownr += 1;
                    row = wb.getSheetAt(i + count).createRow(rownr);
                    cell = row.createCell(0);
                    cell.setCellValue(auth_repr[i]);
                    cell.setCellStyle(normal);
                    cell = row.createCell(3);
                    String nameManager = null;
                    if (signComboBox.getSelectedIndex() == 0) {
                        nameManager = "José Maas, Quality Manager";
                    } else {
                        nameManager = "Ad Netten, Quality Manager";
                    }
                    cell.setCellValue(nameManager);
                    cell.setCellStyle(normal);

                    rownr += 1;
                    String stamp = null;
                    if (signComboBox.getSelectedIndex() == 0) {
                        stamp = "signJose.png";
                    } else {
                        stamp = "signAd.png";
                    }
                    FileInputStream signManager = new FileInputStream(images + "\\" + stamp);

                    byte[] signManager_bytes = IOUtils.toByteArray(signManager);
                    int signManager_id = wb.addPicture(signManager_bytes, XSSFWorkbook.PICTURE_TYPE_PNG);
                    signManager.close();
                    XSSFDrawing signManager_drawing = wb.getSheetAt(i + count).createDrawingPatriarch();
                    XSSFClientAnchor signManager_anchor = new XSSFClientAnchor();
                    signManager_anchor.setCol1(0);
                    signManager_anchor.setRow1(rownr);
                    XSSFPicture signManager_picture = signManager_drawing.createPicture(signManager_anchor, signManager_id);
                    signManager_picture.resize();

                    rownr += 5;
                    row = wb.getSheetAt(i + count).createRow(rownr);
                    cell = row.createCell(0);
                    cell.setCellValue(place[i] + DateFormat.format(Date).toString());
                    cell.setCellStyle(normal);

                    FileInputStream CE = new FileInputStream(images + "\\CE.jpg");
                    byte[] CE_bytes = IOUtils.toByteArray(CE);
                    int CE_id = wb.addPicture(CE_bytes, XSSFWorkbook.PICTURE_TYPE_JPEG);
                    CE.close();
                    XSSFDrawing CE_drawing = wb.getSheetAt(i + count).createDrawingPatriarch();
                    XSSFClientAnchor CE_anchor = new XSSFClientAnchor();
                    CE_anchor.setCol1(5);
                    CE_anchor.setRow1(rownr - 1);
                    XSSFPicture CE_picture = CE_drawing.createPicture(CE_anchor, CE_id);
                    CE_picture.resize(0.2);

                    if (rf_f.length() > 0) {
                        row = wb.getSheetAt(i + count).createRow(rownr - 4);
                        cell = row.createCell(7);
                        cell.setCellValue("R&TTE APPROVED");
                        cell.setCellStyle(normalHW);
                        wb.getSheetAt(i + count).addMergedRegion(new CellRangeAddress(rownr - 4, rownr - 3, 7, 8));

                        row = wb.getSheetAt(i + count).createRow(rownr - 1);
                        cell = row.createCell(7);
                        cell.setCellValue(rf_f);
                        cell.setCellStyle(normalHW);
                        wb.getSheetAt(i + count).addMergedRegion(new CellRangeAddress(rownr - 1, rownr, 7, 8));
                    }

                    rownr += 2;
                    String pictureName = null;
                    File pic = null;
                    String pictureName2 = productContent + "\\" + sapWithoutDocs + "\\LR_" + sapWithoutDocs + "_2.jpg";
                    String pictureName3 = productContent + "\\" + sapWithoutDocs + "\\LR_" + sapWithoutDocs + "_3.jpg";
                    String pictureName4 = productContent + "\\" + sapWithoutDocs + "\\LR_" + sapWithoutDocs + "_4.jpg";
                    String pictureName10 = productContent + "\\" + sapWithoutDocs + "\\LR_" + sapWithoutDocs + "_10.jpg";
                    File pic2 = new File(pictureName2);
                    File pic3 = new File(pictureName3);
                    File pic4 = new File(pictureName4);
                    File pic10 = new File(pictureName10);
                    if (pic2.exists()) {
                        pictureName = pictureName2;
                        pic = pic2;
                    } else if (pic3.exists()) {
                        pictureName = pictureName3;
                        pic = pic3;
                    } else if (pic4.exists()) {
                        pictureName = pictureName4;
                        pic = pic4;
                    } else if (pic10.exists()) {
                        pictureName = pictureName10;
                        pic = pic10;
                    } else {
                        pictureName = null;
                        pic = null;
                    }

                    if (pictureName != null) {
                        BufferedImage bufferedImage = ImageIO.read(pic);
                        File pic_new = new File(docFolder + "\\" + pic.getName());
                        ImageIO.write(bufferedImage, "jpg", pic_new);
                        FileInputStream pict = new FileInputStream(pic_new);
                        byte[] pict_bytes = IOUtils.toByteArray(pict);
                        int pict_id = wb.addPicture(pict_bytes, XSSFWorkbook.PICTURE_TYPE_JPEG);
                        pict.close();
                        XSSFDrawing pict_drawing = wb.getSheetAt(i + count).createDrawingPatriarch();
                        XSSFClientAnchor pict_anchor = new XSSFClientAnchor();
                        pict_anchor.setCol1(1);
                        pict_anchor.setRow1(rownr);
                        XSSFPicture pict_picture = pict_drawing.createPicture(pict_anchor, pict_id);
                        pict_picture.resize(0.3);
                        pic_new.delete();

                        rownr = pict_anchor.getRow2() + 1;
                    }
                    wb.getSheetAt(i + count).setFitToPage(true);
                    wb.getSheetAt(i + count).setAutobreaks(true);
                    PrintSetup ps = wb.getSheetAt(i + count).getPrintSetup();
                    ps.setFitHeight((short) 1);
                    ps.setFitWidth((short) 1);
                    ps.setPaperSize(PrintSetup.A4_PAPERSIZE);
                }

            }
        } catch (Exception X) {
        }
    }

    private void createIntro(Date Date, String sapNo, XSSFWorkbook wb) {
        System.out.println("Set: " + sapNo);
        DateFormat DateFormat = new SimpleDateFormat("dd-MM-yyyy");
        String sapWithoutDocs = sapNo.replace(".", "");

        CellStyle header = wb.createCellStyle();
        XSSFFont arialBlack = wb.createFont();
        arialBlack.setFontName("Arial Black");
        arialBlack.setFontHeightInPoints((short) 12);
        arialBlack.setBold(true);
        header.setFont(arialBlack);
        header.setAlignment(XSSFCellStyle.ALIGN_CENTER);

        CellStyle itemStyle = wb.createCellStyle();
        itemStyle.setFont(arialBlack);
        itemStyle.setAlignment(XSSFCellStyle.ALIGN_LEFT);

        CellStyle normal = wb.createCellStyle();
        XSSFFont arial = wb.createFont();
        arial.setFontName("Arial");
        arial.setFontHeightInPoints((short) 10);
        normal.setFont(arial);

        CellStyle normalW = wb.createCellStyle();
        normalW.setFont(arial);
        normalW.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        normalW.setWrapText(true);

        CellStyle normalH = wb.createCellStyle();
        XSSFFont arialH = wb.createFont();
        arialH.setFontName("Arial");
        arialH.setFontHeightInPoints((short) 12);
        normalH.setFont(arialH);

        CellStyle normalHI = wb.createCellStyle();
        XSSFFont arialHI = wb.createFont();
        arialHI.setFontName("Arial");
        arialHI.setItalic(true);
        arialHI.setFontHeightInPoints((short) 12);
        normalHI.setFont(arialHI);

        CellStyle normalB = wb.createCellStyle();
        XSSFFont arialB = wb.createFont();
        arialB.setFontName("Arial Black");
        arialB.setFontHeightInPoints((short) 10);
        arialB.setBold(true);
        normalB.setFont(arialB);

        CellStyle normalHW = wb.createCellStyle();
        arialH.setBold(true);
        normalHW.setFont(arialH);
        normalHW.setWrapText(true);
        normalHW.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        normalHW.setAlignment(XSSFCellStyle.ALIGN_CENTER);

        String sql = "SELECT "
                + "SAP,ITEM,BRAND,DESCR_EN,DESCR_FR,DESCR_DE"
                + " FROM elro.items"
                + " WHERE SAP='" + sapNo + "';";
        con = Utils.getConnection();

        try {
            st = con.createStatement();
            rs = st.executeQuery(sql);

            while (rs.next()) {
                sap = rs.getString(1);
                item = rs.getString(2);
                brand = rs.getString(3);
                descr_en = rs.getString(4);
                descr_fr = rs.getString(5);
                descr_de = rs.getString(6);

                String[] declaration = {"EU DECLARATION OF CONFORMITY", "EU-KONFORMITÄTSERKLÄRUNG", "DÉCLARATION UE DE CONFORMITÉ"};
                String[] company = {"Company:", "Firma:", "Société:"};
                String[] address = {"Address, City:", "Addresse:", "Adresse/Ville:"};
                String[] country1 = {"Country:", "Land:", "Pays:"};
                String[] country2 = {"The Netherlands", "Niederlande", "Pays-Bas"};
                String[] object = {"Object of the declaration:", "Gegenstand der Erklärung:", "Objet de la déclaration:"};
                String[] descr = {"Description:", "Artikelbeschreibung:", "Description:"};
                String[] descr_lang = {descr_en, descr_de, descr_fr};
                String[] name = {"Pruduct name:", "Artikel-Nr.:", "Référence produit:"};
                String[] trademark = {"Trademark:", "Markenname:", "Marque déposée:"};
                String[] auth_repr = {"Authorized representative:", "Bevollmächtigten Vertreter:", "Représentant autorisé:"};
                String[] place = {"Place and date of issue: Gilze, ", "Ort und Datum der Ausstellung: Gilze, ", "Lieu et date d’établissement: Gilze, "};

                int rownr = 0;
                XSSFRow row = wb.getSheetAt(0).createRow(rownr);//0
                XSSFCell cell = row.createCell(0);
                cell.setCellValue(declaration[0]);
                cell.setCellStyle(header);
                wb.getSheetAt(0).addMergedRegion(new CellRangeAddress(rownr, rownr, 0, 8));

                rownr += 2;
                row = wb.getSheetAt(0).createRow(rownr);//2
                cell = row.createCell(0);
                cell.setCellValue(company[0]);
                cell.setCellStyle(normal);
                cell = row.createCell(2);
                cell.setCellValue("Smartwares Safety & Lighting B.V.");
                cell.setCellStyle(normal);

                FileInputStream logo = new FileInputStream(images + "\\logo.png");
                byte[] logo_bytes = IOUtils.toByteArray(logo);
                int logo_id = wb.addPicture(logo_bytes, XSSFWorkbook.PICTURE_TYPE_PNG);
                logo.close();
                XSSFDrawing logo_drawing = wb.getSheetAt(0).createDrawingPatriarch();
                XSSFClientAnchor logo_anchor = new XSSFClientAnchor();
                logo_anchor.setCol1(6);
                logo_anchor.setRow1(rownr);
                logo_anchor.setDx1(0);
                logo_anchor.setDy1(0);
                logo_anchor.setCol2((short) 8);
                logo_anchor.setRow2(rownr + 2);
                logo_anchor.setDx2(0);
                logo_anchor.setDy2(0);
                XSSFPicture logo_picture = logo_drawing.createPicture(logo_anchor, logo_id);
                logo_picture.resize(0.78);

                rownr += 1;
                row = wb.getSheetAt(0).createRow(rownr);//3
                cell = row.createCell(0);
                cell.setCellValue(address[0]);
                cell.setCellStyle(normal);
                cell = row.createCell(2);
                cell.setCellValue("Broekakkerweg 15, 5126 BD Gilze");
                cell.setCellStyle(normal);

                rownr += 1;
                row = wb.getSheetAt(0).createRow(rownr);//4
                cell = row.createCell(0);
                cell.setCellValue(country1[0]);
                cell.setCellStyle(normal);
                cell = row.createCell(2);
                cell.setCellValue(country2[0]);
                cell.setCellStyle(normal);

                rownr += 2;
                row = wb.getSheetAt(0).createRow(rownr);//8
                cell = row.createCell(0);
                cell.setCellValue(object[0]);
                cell.setCellStyle(normal);

                rownr += 2;
                row = wb.getSheetAt(0).createRow(rownr);//10
                cell = row.createCell(0);
                cell.setCellValue(descr[0]);
                cell.setCellStyle(normal);
                cell = row.createCell(2);
                cell.setCellValue(descr_lang[0]);
                cell.setCellStyle(normalH);

                rownr += 1;
                row = wb.getSheetAt(0).createRow(rownr);//11
                cell = row.createCell(0);
                cell.setCellValue(name[0]);
                cell.setCellStyle(normal);
                cell = row.createCell(2);
                cell.setCellValue(item);
                cell.setCellStyle(itemStyle);

                rownr += 1;
                row = wb.getSheetAt(0).createRow(rownr);//12
                cell = row.createCell(0);
                cell.setCellValue(trademark[0]);
                cell.setCellStyle(normal);
                cell = row.createCell(2);
                cell.setCellValue(brand.replace("PL ", ""));
                cell.setCellStyle(normalHI);

                FileInputStream brand_logo = new FileInputStream(images + "\\BRANDS\\" + brand + ".png");
                byte[] brand_logo_bytes = IOUtils.toByteArray(brand_logo);
                int brand_logo_id = wb.addPicture(brand_logo_bytes, XSSFWorkbook.PICTURE_TYPE_PNG);
                brand_logo.close();
                XSSFDrawing brand_logo_drawing = wb.getSheetAt(0).createDrawingPatriarch();
                XSSFClientAnchor brand_logo_anchor = new XSSFClientAnchor();
                brand_logo_anchor.setCol1(5);
                brand_logo_anchor.setRow1(rownr - 1);
                XSSFPicture brand_logo_picture = brand_logo_drawing.createPicture(brand_logo_anchor, brand_logo_id);
                brand_logo_picture.resize();

                rownr += 2;
                row = wb.getSheetAt(0).createRow(rownr);
                cell = row.createCell(0);
                cell.setCellValue("This set contains components mentioned on attached declarations");
                cell.setCellStyle(normal);

                rownr += 2;
                String pictureName = null;
                File pic = null;
                String pictureName2 = productContent + "\\" + sapWithoutDocs + "\\LR_" + sapWithoutDocs + "_2.jpg";
                String pictureName3 = productContent + "\\" + sapWithoutDocs + "\\LR_" + sapWithoutDocs + "_3.jpg";
                String pictureName4 = productContent + "\\" + sapWithoutDocs + "\\LR_" + sapWithoutDocs + "_4.jpg";
                String pictureName10 = productContent + "\\" + sapWithoutDocs + "\\LR_" + sapWithoutDocs + "_10.jpg";
                String pictureName11 = productContent + "\\" + sapWithoutDocs + "\\LR_" + sapWithoutDocs + "_11.jpg";
                String pictureName12 = productContent + "\\" + sapWithoutDocs + "\\LR_" + sapWithoutDocs + "_12.jpg";
                File pic2 = new File(pictureName2);
                File pic3 = new File(pictureName3);
                File pic4 = new File(pictureName4);
                File pic10 = new File(pictureName10);
                File pic11 = new File(pictureName11);
                File pic12 = new File(pictureName12);
                if (pic10.exists()) {
                    pictureName = pictureName2;
                    pic = pic10;
                } else if (pic11.exists()) {
                    pictureName = pictureName3;
                    pic = pic11;
                } else if (pic12.exists()) {
                    pictureName = pictureName4;
                    pic = pic12;
                } else if (pic2.exists()) {
                    pictureName = pictureName10;
                    pic = pic2;
                } else if (pic3.exists()) {
                    pictureName = pictureName10;
                    pic = pic3;
                } else if (pic4.exists()) {
                    pictureName = pictureName10;
                    pic = pic4;
                } else {
                    pictureName = null;
                    pic = null;
                }
                System.out.println(pic.getName());
                if (pictureName != null) {
                    BufferedImage bufferedImage = ImageIO.read(pic);
                    File pic_new = new File(docFolder + "\\" + pic.getName());
                    ImageIO.write(bufferedImage, "jpg", pic_new);
                    FileInputStream pict = new FileInputStream(pic_new);
                    byte[] pict_bytes = IOUtils.toByteArray(pict);
                    int pict_id = wb.addPicture(pict_bytes, XSSFWorkbook.PICTURE_TYPE_JPEG);
                    pict.close();
                    XSSFDrawing pict_drawing = wb.getSheetAt(0).createDrawingPatriarch();
                    XSSFClientAnchor pict_anchor = new XSSFClientAnchor();
                    pict_anchor.setCol1(1);
                    pict_anchor.setRow1(rownr);
                    XSSFPicture pict_picture = pict_drawing.createPicture(pict_anchor, pict_id);
                    pict_picture.resize(0.3);
                    pic_new.delete();

                    rownr = pict_anchor.getRow2() + 1;
                }
                wb.getSheetAt(0).setFitToPage(true);
                wb.getSheetAt(0).setAutobreaks(true);
                PrintSetup ps = wb.getSheetAt(0).getPrintSetup();
                ps.setFitHeight((short) 1);
                ps.setFitWidth((short) 1);
                ps.setPaperSize(PrintSetup.A4_PAPERSIZE);

            }
        } catch (Exception X) {
        }
    }

    private void createPDF() {
        if (Desktop.isDesktopSupported()) {
            desktop = Desktop.getDesktop();
        }
        try {
            desktop.open(new File("H:\\DoCs_temporary\\MakePDF\\MakePDF.xlsm"));
        } catch (IOException ex) {
            Logger.getLogger(DoC.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
}
