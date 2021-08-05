package zcrb.excel.merging;

import java.awt.Color;
import java.awt.Component;
import java.awt.Cursor;
import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.font.TextAttribute;
import java.io.File;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.ResourceBundle;

import javax.swing.Box;
import javax.swing.BoxLayout;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.SwingUtilities;
import javax.swing.WindowConstants;

import zcrb.excel.merging.utils.UtfResourceBundle;

public class MainFrame extends JFrame {

  private static final long serialVersionUID = -1811523675073957560L;
  final private Model model;
  final private UtfResourceBundle bundle;
  final private String HELP_URL = "https://github.com/andrejs82git/merge-excel-files";

  public MainFrame() {
    super();
    bundle = new UtfResourceBundle(ResourceBundle.getBundle("bundle"));
    setTitle(bundle.getString("MainFrame.title.text"));
    this.model = new Model();
    this.setContentPane(createInterface());
    this.setSize(800, 400);

    this.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
    this.setVisible(true);
  }

  private JPanel createInterface() {
    JPanel mainPanel = new JPanel();
    mainPanel.setLayout(new GridBagLayout());
    GridBagConstraints c = new GridBagConstraints();
    c.insets = new Insets(2, 2, 2, 2);

    {
      JPanel firstPanel = new JPanel();
      firstPanel.setLayout(new BoxLayout(firstPanel, BoxLayout.LINE_AXIS));
      firstPanel.add(this.getChooseDirectoryButton());
      firstPanel.add(Box.createRigidArea(new Dimension(5, 0)));
      firstPanel.add(this.getFileLabel());
      firstPanel.add(Box.createRigidArea(new Dimension(5, 0)));
      firstPanel.add(Box.createHorizontalGlue());
      firstPanel.add(this.getHelpButton());

      c.gridx = 0;
      c.gridy = 0;
      c.fill = GridBagConstraints.HORIZONTAL;
      c.anchor = GridBagConstraints.NORTHWEST;
      mainPanel.add(firstPanel, c);
    }

    {
      JPanel secondPanel = new JPanel();
      secondPanel.setLayout(new BoxLayout(secondPanel, BoxLayout.LINE_AXIS));
      JLabel statusLabel = getTemplateJLabel();
      secondPanel.add(statusLabel);

      c.gridx = 0;
      c.gridy = 1;
      c.fill = GridBagConstraints.HORIZONTAL;
      c.anchor = GridBagConstraints.NORTHWEST;
      mainPanel.add(secondPanel, c);
    }

    {
      JPanel thirdPanel = new JPanel();
      thirdPanel.setLayout(new BoxLayout(thirdPanel, BoxLayout.LINE_AXIS));
      JCheckBox debugCheckBox = createDebugCheckBox();
      thirdPanel.add(debugCheckBox);

      c.gridx = 0;
      c.gridy = 2;
      c.anchor = GridBagConstraints.NORTHWEST;
      mainPanel.add(thirdPanel, c);
    }

    {
      c.gridx = 0;
      c.gridy = 3;
      c.fill = GridBagConstraints.BOTH;
      c.weightx = 1;
      c.weighty = 1;
      mainPanel.add(this.getListOfFiles(), c);
    }

    {
      c.gridx = 0;
      c.gridy = 9;
      c.fill = GridBagConstraints.BOTH;
      c.weightx = 0;
      c.weighty = 1;
      mainPanel.add(this.getMsgTextArea(), c);
    }

    {
      c.gridx = 0;
      c.gridy = 10;
      c.fill = GridBagConstraints.BOTH;
      c.weightx = 1;
      c.weighty = 0;
      mainPanel.add(this.getRunButton(), c);
    }

    // DragAndDrop
    new java.awt.dnd.DropTarget(mainPanel, new MyDragDropListener() {
      @Override
      public void doDragFolder(File f) {
        MainFrame.this.model.setSrcDir(f);
      }
    });

    return mainPanel;
  }

  private Component getMsgTextArea() {
    JTextArea textArea = new JTextArea();
    SimpleDateFormat formatD = new SimpleDateFormat("[dd.mm.yyyy HH:mm:ss]");
    MainFrame.this.model.addPropertyChangeListener(Model.MODEL_SEND_MESSAGE,
        (evt) -> textArea.append(formatD.format(new Date()) + " " + evt.getNewValue() + "\n"));

    JScrollPane scroll = new JScrollPane(textArea);
    scroll.setMinimumSize(new Dimension(-1, 70));
    // scroll.setMaximumSize(new Dimension(-1, 1000));
    // scroll.setSize(new Dimension(-1, 70));
    return scroll;
  }

  private Component getRunButton() {
    JButton jButtonStartAddExcelFiles = new JButton(bundle.getString("MainFrame.jButtonStartAddExcelFiles.text"));
    jButtonStartAddExcelFiles.setEnabled(false);
    this.model.addPropertyChangeListener(Model.CAN_BE_RUN_CHANGED, evt -> {
      jButtonStartAddExcelFiles.setEnabled((Boolean) evt.getNewValue());
    });

    jButtonStartAddExcelFiles.addActionListener(evt -> {
      try {
        PoiMerging.runAdder(MainFrame.this.model);
      } catch (Exception e) {
        e.printStackTrace();
        StringWriter sw = new StringWriter();
        PrintWriter pw = new PrintWriter(sw);
        e.printStackTrace(pw);
        MainFrame.this.model.sendMsg(bundle.getString("MainFrame.ErrorWhileProcessMerge.text") + ":\n" + e.getMessage()
            + "\n" + sw.toString() + "\n");
      }
    });
    return jButtonStartAddExcelFiles;
  }

  private JCheckBox createDebugCheckBox() {
    JCheckBox jCheckBoxDebugMode = new JCheckBox(bundle.getString("MainFrame.jCheckBoxDebugMode.text"));
    jCheckBoxDebugMode.addActionListener(ev -> MainFrame.this.model.setDebugMode(jCheckBoxDebugMode.isSelected()));
    return jCheckBoxDebugMode;
  }

  private JLabel getTemplateJLabel() {
    JLabel jLabelTemplate = new JLabel("-");
    MainFrame.this.model.addPropertyChangeListener(Model.TEMPLATE_FILE_CHANGED, evt -> {
      File f = (File) evt.getNewValue();
      if (f == null) {
        jLabelTemplate.setText(bundle.getString("MainFrame.jLabelTemplate.FileNotFound.text"));
        jLabelTemplate.setForeground(Color.RED);
      } else {
        jLabelTemplate
            .setText(bundle.getString("MainFrame.jLabelTemplate.FileFound.text") + ": " + f.getAbsolutePath());
        jLabelTemplate.setForeground(Color.decode("#006e00"));
      }
    });
    return jLabelTemplate;
  }

  private Component getFileLabel() {
    JLabel jLabelDirectory = new JLabel("-");
    MainFrame.this.model.addPropertyChangeListener(Model.CHANGE_SRC_DIR, ev -> {
      File f = (File) ev.getNewValue();
      jLabelDirectory.setText(bundle.getString("MainFrame.jLabelDirectory.text") + ": " + f.getAbsolutePath());
    });

    jLabelDirectory.setCursor(new Cursor(Cursor.HAND_CURSOR));
    jLabelDirectory.addMouseListener(new MouseAdapter() {
      @Override
      public void mouseClicked(MouseEvent e) {
        if (SwingUtilities.isLeftMouseButton(e)) {
          Desktop desktop = Desktop.getDesktop();
          try {
            File srcDir = MainFrame.this.model.getSrcDir();
            if (srcDir != null) {
              desktop.open(srcDir);
            }
          } catch (Exception err) {
            err.printStackTrace();
          }
        }
      }

      Font standartFont = jLabelDirectory.getFont();

      public void mouseEntered(MouseEvent e) {
        Map<TextAttribute, Object> attributes = new HashMap<>(standartFont.getAttributes());
        attributes.put(TextAttribute.UNDERLINE, TextAttribute.UNDERLINE_ON);
        jLabelDirectory.setFont(new Font(attributes));
      }

      public void mouseExited(MouseEvent e) {
        jLabelDirectory.setFont(standartFont);
      }
    });
    return jLabelDirectory;
  }

  private Component getHelpButton() {
    JButton helpButton = new JButton(bundle.getString("MainFrame.jHelpButton.text"));
    helpButton.addActionListener(new ActionListener() {
      public void actionPerformed(ActionEvent e) {
        try {
          Desktop.getDesktop().browse(new URL(HELP_URL).toURI());
        } catch (Exception ex) {
          ex.printStackTrace();
        }
      }
    });
    return helpButton;
  }

  private Component getListOfFiles() {
    FilesTableModel tableModel = new FilesTableModel();
    this.model.addPropertyChangeListener(Model.EXCEL_FILES_CHANGED, evt -> {
      File[] files = (File[]) evt.getNewValue();
      tableModel.setNewData(Arrays.asList(files));
    });

    JTable table = new JTable(tableModel);
    return new JScrollPane(table);
  }

  private JButton getChooseDirectoryButton() {
    JButton jButtonChooseDirectory = new JButton(bundle.getString("MainFrame.jButtonChooseDirectory.text"));
    jButtonChooseDirectory.addActionListener(new ActionListener() {

      @Override
      public void actionPerformed(ActionEvent e) {
        JFileChooser chooser = new JFileChooser();
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

        int returnVal = chooser.showOpenDialog(MainFrame.this);
        if (returnVal == JFileChooser.APPROVE_OPTION) {
          MainFrame.this.model.setSrcDir(chooser.getSelectedFile());
        }
      }
    });

    return jButtonChooseDirectory;
  }
}
