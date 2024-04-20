package com.gdsc.testing;

import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.swing.*;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


import org.apache.poi.ss.usermodel.*;

/**
 * Hello world!
 *
 */
public class App extends JFrame implements ActionListener
{

    private JButton btnSelectFile;
    private JLabel lblMessage, lblWelcome, lblDescription;


    public App() {
        super("Excel Step Converter");

        // Welcome message and description labels
        lblWelcome = new JLabel("Welcome to the Step Converter!", JLabel.CENTER);
        lblWelcome.setFont(new Font("Arial", Font.BOLD, 16));
        lblDescription = new JLabel("This tool converts XML steps defined to clear, formatted text.", JLabel.CENTER);

        btnSelectFile = new JButton("Select Excel File");
        btnSelectFile.addActionListener(this);

        lblMessage = new JLabel("");

        // Set color scheme for a user-friendly experience
        Color backgroundColor = new Color(230, 245, 255); // Light blue for background
        Color textColor = new Color(50, 50, 50); // Dark gray for text
        Color buttonColor = new Color(100, 149, 237); // Light blue button

        getContentPane().setBackground(backgroundColor); // Set background color for entire window
        lblWelcome.setForeground(textColor);
        lblDescription.setForeground(textColor);
        btnSelectFile.setBackground(buttonColor);
        btnSelectFile.setForeground(textColor);

        // Arrange components and set layout
        JPanel panel = new JPanel(new GridLayout(4, 1, 5, 5));
        panel.setBackground(backgroundColor); // Set background color for the panel
        panel.add(lblWelcome);
        panel.add(lblDescription);
        panel.add(btnSelectFile);
        panel.add(lblMessage);

        add(panel);

        setSize(500, 200);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setVisible(true);
    }

    @Override
    public void actionPerformed(ActionEvent e) {
        if (e.getSource() == btnSelectFile) {
            JFileChooser fileChooser = new JFileChooser();
            int selection = fileChooser.showOpenDialog(this);
            if (selection == JFileChooser.APPROVE_OPTION) {
                String filePath = fileChooser.getSelectedFile().getAbsolutePath();
                try {
                    convertExcelFile(filePath);
                    lblMessage.setText("Successfully processed and modified Excel file!");
                } catch (Exception ex) {
                    ex.printStackTrace();
                    lblMessage.setText("Error processing file: " + ex.getMessage());
                }
            }
        }
    }

    public void convertExcelFile(String filePath) throws Exception {
        java.util.List<String> stepsData = readStepsFromCSV(filePath);

        java.util.List<String> processedSteps = new ArrayList<>();
        for (String step : stepsData) {
            processedSteps.add(processStep(step));
        }

        modifyExcelFile(filePath, processedSteps);

        System.out.println("Successfully processed steps and modified existing Excel file: " + filePath);
    }

    private java.util.List<String> readStepsFromCSV(String csvPath) throws IOException {
        java.util.List<String> stepsData = new ArrayList<>();
        try (FileInputStream inputStream = new FileInputStream(csvPath)) {
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0);

            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    continue;
                }
                Cell cell = row.getCell(2);  // Assuming "steps" is in the third column (index 2)
                if (cell != null) {
                    String stepValue = cell.getStringCellValue().trim();
                    stepsData.add(stepValue);
                }
            }
        }
        return stepsData;
    }

    private String processStep(String step) throws Exception {
        String xmlString = step;

        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        DocumentBuilder builder = factory.newDocumentBuilder();
        Document document = builder.parse(new ByteArrayInputStream(xmlString.getBytes()));

        //write here column name from old exported sheet
        NodeList steps = document.getElementsByTagName("step");

        String processedStep = "";
        for (int i = 0; i < steps.getLength(); i++) {
            Node stepNode = steps.item(i);


            String text = getChildText(stepNode, "parameterizedString", 0);
            text = removeHtmlTags(text);

            String expected = getChildText(stepNode, "parameterizedString", 1);
            expected = removeHtmlTags(expected);


            processedStep += "Step " + (i + 1) + ": " + text + "\nExpected: " + expected + "\n";
        }

        return processedStep.trim();
    }

    private static String getChildText(Node parentNode, String childName, int index) {
        NodeList childNodes = parentNode.getChildNodes();
        for (int i = 0; i < childNodes.getLength(); i++) {
            Node childNode = childNodes.item(i);
            if (childNode.getNodeType() == Node.ELEMENT_NODE && childNode.getNodeName().equals(childName)) {
                if (index == 0) {
                    return childNode.getTextContent().trim();
                }
                index--;
            }
        }
        return "";
    }

    private static String removeHtmlTags(String text) {
        Pattern pattern = Pattern.compile("<[^>]+>");  // Matches any HTML tag
        Matcher matcher = pattern.matcher(text);
        return matcher.replaceAll("").trim();
    }

    private static void modifyExcelFile(String csvPath, List<String> processedSteps) throws IOException {
        Workbook workbook;
        try (FileInputStream inputStream = new FileInputStream(csvPath)) {
            workbook = WorkbookFactory.create(inputStream);
        }

        Sheet sheet = workbook.getSheetAt(0);

        // Assuming "steps" is in the third column (index 2)
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                continue;
            }
            Cell cell = row.getCell(2);
            if (cell == null) {
                cell = row.createCell(2);  // Create the cell if it doesn't exist
            }
            cell.setCellValue(processedSteps.get(rowIndex - 1));  // Update cell value with processed step
        }

        try (FileOutputStream outputStream = new FileOutputStream(csvPath)) {
            workbook.write(outputStream);
        }
    }



    public static void main(String[] args) {

        new App();

    }
}