package excelConvertor;

import javax.swing.*;
import javax.swing.border.*;
import java.awt.*;
import java.awt.event.*;

public class MainApp extends JFrame {
    private static final Color PRIMARY_COLOR = new Color(51, 153, 255);
    private static final Color BACKGROUND_COLOR = new Color(240, 240, 245);
    private static final Font HEADER_FONT = new Font("Segue UI", Font.BOLD, 24);
    private static final Font LABEL_FONT = new Font("Segue UI", Font.BOLD, 14);
    private static final Font FIELD_FONT = new Font("Segue UI", Font.PLAIN, 14);

    public MainApp() {
        setTitle("MTH Excel to JSON Converter");
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setSize(800, 600);
        setLocationRelativeTo(null);
        setMinimumSize(new Dimension(600, 500));

        // Create main panel with padding
        JPanel mainPanel = new JPanel(new BorderLayout(20, 20));
        mainPanel.setBackground(BACKGROUND_COLOR);
        mainPanel.setBorder(new EmptyBorder(30, 40, 30, 40));

        // Add header
        JLabel headerLabel = new JLabel("MTH Excel to JSON Converter");
        headerLabel.setFont(HEADER_FONT);
        headerLabel.setForeground(PRIMARY_COLOR);
        headerLabel.setHorizontalAlignment(SwingConstants.CENTER);
        mainPanel.add(headerLabel, BorderLayout.NORTH);

        // Create form panel
        JPanel formPanel = new JPanel(new GridBagLayout());
        formPanel.setBackground(BACKGROUND_COLOR);
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(10, 5, 10, 5);

        // Initialize components with improved styling
        JTextField filePathField = createStyledTextField();
        JTextField sheetNameField = createPlaceholderTextField();
        JTextField columnsField = createStyledTextField();
        JTextField outputLocationField = createStyledTextField();

        JButton browseFileButton = createStyledButton("Browse");
        JButton browseOutputButton = createStyledButton("Browse");
        JButton convertButton = createPrimaryButton();

        // Add components to form panel
        addFormRow(formPanel, gbc, 0, "Excel File:", filePathField, browseFileButton);
        addFormRow(formPanel, gbc, 1, "Sheet Name:", sheetNameField, null);
        addFormRow(formPanel, gbc, 2, "Columns:", columnsField,
                createInfoLabel());
        addFormRow(formPanel, gbc, 3, "Output Location:", outputLocationField, browseOutputButton);

        // Add convert button with special styling
        gbc.gridy = 4;
        gbc.gridx = 0;
        gbc.gridwidth = 3;
        gbc.insets = new Insets(30, 5, 10, 5);
        gbc.anchor = GridBagConstraints.CENTER;
        formPanel.add(convertButton, gbc);

        // Add form panel to main panel
        mainPanel.add(formPanel, BorderLayout.CENTER);

        // Add status panel at the bottom
        JPanel statusPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        statusPanel.setBackground(BACKGROUND_COLOR);
        JLabel statusLabel = new JLabel("Ready");
        statusLabel.setFont(FIELD_FONT);
        statusPanel.add(statusLabel);
        mainPanel.add(statusPanel, BorderLayout.SOUTH);

        setContentPane(mainPanel);

        // Add action listeners
        browseFileButton.addActionListener(e -> browsePath(filePathField, JFileChooser.FILES_ONLY,
                "Select Excel File"));

        browseOutputButton.addActionListener(e -> browsePath(outputLocationField,
                JFileChooser.DIRECTORIES_ONLY, "Select Output Directory"));

        convertButton.addActionListener(e -> {
            try {
                performConversion(filePathField.getText(), sheetNameField.getText(),
                        columnsField.getText(), outputLocationField.getText(), statusLabel);
            } catch (Exception ex) {
                showError("Conversion Error", ex.getMessage());
            }
        });
    }

    private JTextField createStyledTextField() {
        JTextField field = new JTextField(20);
        field.setFont(FIELD_FONT);
        field.setMargin(new Insets(8, 10, 8, 10));
        return field;
    }

    private JTextField createPlaceholderTextField() {
        JTextField field = createStyledTextField();
        field.setForeground(Color.GRAY);
        field.setText("Sheet1 (default)");

        field.addFocusListener(new FocusAdapter() {
            @Override
            public void focusGained(FocusEvent e) {
                if (field.getText().equals("Sheet1 (default)")) {
                    field.setText("");
                    field.setForeground(Color.BLACK);
                }
            }

            @Override
            public void focusLost(FocusEvent e) {
                if (field.getText().isEmpty()) {
                    field.setText("Sheet1 (default)");
                    field.setForeground(Color.GRAY);
                }
            }
        });

        return field;
    }

    private JButton createStyledButton(String text) {
        JButton button = new JButton(text);
        button.setFont(FIELD_FONT);
        button.setFocusPainted(false);
        button.setMargin(new Insets(8, 15, 8, 15));
        return button;
    }

    private JButton createPrimaryButton() {
        JButton button = createStyledButton("Convert to JSON");
        button.setBackground(PRIMARY_COLOR);
        button.setForeground(Color.black);
        button.setFont(new Font("Segue UI", Font.BOLD, 16));
        button.setBorder(new EmptyBorder(12, 30, 12, 30));

        button.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseEntered(MouseEvent e) {
                button.setBackground(PRIMARY_COLOR.darker());
            }

            @Override
            public void mouseExited(MouseEvent e) {
                button.setBackground(PRIMARY_COLOR);
            }
        });

        return button;
    }

    private JLabel createInfoLabel() {
        JLabel label = new JLabel("Enter comma-separated column numbers");
        label.setFont(new Font("Segue UI", Font.ITALIC, 12));
        label.setForeground(Color.GRAY);
        return label;
    }

    private void addFormRow(JPanel panel, GridBagConstraints gbc, int row,
                            String labelText, JComponent field, JComponent extra) {
        gbc.gridy = row;

        // Add label
        gbc.gridx = 0;
        gbc.gridwidth = 1;
        gbc.weightx = 0.0;
        JLabel label = new JLabel(labelText);
        label.setFont(LABEL_FONT);
        panel.add(label, gbc);

        // Add field
        gbc.gridx = 1;
        gbc.weightx = 1.0;
        panel.add(field, gbc);

        // Add extra component if provided
        if (extra != null) {
            gbc.gridx = 2;
            gbc.weightx = 0.0;
            panel.add(extra, gbc);
        }
    }

    private void browsePath(JTextField field, int mode, String title) {
        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle(title);
        chooser.setFileSelectionMode(mode);

        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            field.setText(chooser.getSelectedFile().getAbsolutePath());
        }
    }

    private void performConversion(String filePath, String sheetName,
                                   String columnsText, String outputLocation, JLabel statusLabel) {
        // Validate inputs
        if (filePath.isEmpty() || columnsText.isEmpty() || outputLocation.isEmpty()) {
            throw new IllegalArgumentException("All fields are required!");
        }

        // Parse columns
        String[] colsArray = columnsText.split(",");
        int[] cols = new int[colsArray.length];
        try {
            for (int i = 0; i < colsArray.length; i++) {
                cols[i] = Integer.parseInt(colsArray[i].trim());
            }
        } catch (NumberFormatException e) {
            throw new IllegalArgumentException("Invalid column numbers format!");
        }

        // Update status
        statusLabel.setText("Converting...");

        // Perform conversion in background
        SwingWorker<Void, Void> worker = new SwingWorker<>() {
            @Override
            protected Void doInBackground()  {
                ReadExcel readExcel = new ReadExcel();
                String data = readExcel.getExcelData(filePath,
                        sheetName.isEmpty() ? null : sheetName, cols);

                WriteJson writer = new WriteJson(data, outputLocation);
                writer.fileWrite();
                return null;
            }

            @Override
            protected void done() {
                try {
                    get(); // Check for exceptions
                    statusLabel.setText("Conversion completed successfully!");
                    showSuccess();
                } catch (Exception e) {
                    statusLabel.setText("Conversion failed!");
                    showError("Error", "Conversion failed: " + e.getMessage());
                } finally {
                    // Reset status label to "Ready" after conversion
                    statusLabel.setText("Ready");
                }
            }
        };

        worker.execute();
    }

    private void showError(String title, String message) {
        JOptionPane.showMessageDialog(this, message, title,
                JOptionPane.ERROR_MESSAGE);
    }

    private void showSuccess() {
        JOptionPane.showMessageDialog(this, "Conversion completed successfully!", "Success",
                JOptionPane.INFORMATION_MESSAGE);
    }

    public static void main(String[] args) {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception e) {
            e.printStackTrace();
        }

        SwingUtilities.invokeLater(() -> new MainApp().setVisible(true));
    }
}