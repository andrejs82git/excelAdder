package zcrb.excel.adder;

import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.FlowLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;

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
import javax.swing.WindowConstants;
import javax.swing.table.TableColumn;

public class MainFrame extends JFrame {

  private static final long serialVersionUID = -1811523675073957560L;
  final private Model model;

  public MainFrame() {
    super("Склейка excel файлов");
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
      firstPanel.add(this.getChooseButton());
      firstPanel.add(Box.createRigidArea(new Dimension(5, 0)));
      firstPanel.add(this.getFileLabel());

      c.gridx = 0;
      c.gridy = 0;
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

    return mainPanel;
  }

  private Component getMsgTextArea() {
    JTextArea textArea = new JTextArea();
    SimpleDateFormat formatD = new SimpleDateFormat("[dd.mm.yyyy HH:mm:ss]");
    MainFrame.this.model.addPropertyChangeListener(Model.MODEL_SEND_MESSAGE,
        (evt) -> textArea.append(formatD.format(new Date()) + " " + evt.getNewValue()+"\n"));

    JScrollPane scroll = new JScrollPane(textArea);
    scroll.setMinimumSize(new Dimension(-1, 70));
    // scroll.setMaximumSize(new Dimension(-1, 1000));
    // scroll.setSize(new Dimension(-1, 70));
    return scroll;
  }

  private Component getRunButton() {
    JButton button = new JButton("Начать выполнение склейки");
    button.setEnabled(false);
    this.model.addPropertyChangeListener(Model.CAN_BE_RUN_CHANGED, evt -> {
      button.setEnabled((Boolean) evt.getNewValue());
    });

    button.addActionListener(evt -> {
      try {
        PoiAdder.runAdder(MainFrame.this.model);
      } catch (Exception e) {
        e.printStackTrace();
        MainFrame.this.model
            .sendMsg("Произошла ошибка во время выполнения склейки:\n" + e.getMessage() + "\n" + e.getStackTrace());
      }
    });
    return button;
  }

  private JCheckBox createDebugCheckBox() {
    JCheckBox checkBox = new JCheckBox("Добавить в результирующий файл расширенную информацию в комментарии ячеек?");
    checkBox.addActionListener(ev -> MainFrame.this.model.setDebugMode(checkBox.isSelected()));
    return checkBox;
  }

  private JLabel getTemplateJLabel() {
    JLabel label = new JLabel("-");
    MainFrame.this.model.addPropertyChangeListener(Model.TEMPLATE_FILE_CHANGED, evt -> {
      File f = (File) evt.getNewValue();
      if (f == null) {
        label.setText("Файлы template.xls или template.xlsx не найдены");
        label.setForeground(Color.RED);
      } else {
        label.setText("Найден файл шаблона " + f.getAbsolutePath());
        label.setForeground(Color.decode("#006e00"));
      }
    });
    return label;
  }

  private Component getFileLabel() {
    JLabel label = new JLabel("-");
    MainFrame.this.model.addPropertyChangeListener(Model.CHANGE_SRC_DIR, ev -> {
      File f = (File) ev.getNewValue();
      label.setText("Директория: " + f.getAbsolutePath());
    });
    return label;
  }

  private Component getListOfFiles() {
    MyTableModel tableModel = new MyTableModel();
    this.model.addPropertyChangeListener(Model.EXCEL_FILES_CHANGED, evt -> {
      File[] files = (File[]) evt.getNewValue();
      tableModel.setNewData(Arrays.asList(files));
    });

    JTable table = new JTable(tableModel);
    return new JScrollPane(table);
  }

  private JButton getChooseButton() {
    JButton b = new JButton("Выбрать папку");
    b.addActionListener(new ActionListener() {

      @Override
      public void actionPerformed(ActionEvent e) {
        JFileChooser chooser = new JFileChooser("C:\\Users\\andre\\Downloads\\tmp");
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

        int returnVal = chooser.showOpenDialog(MainFrame.this);
        if (returnVal == JFileChooser.APPROVE_OPTION) {
          MainFrame.this.model.setSrcDir(chooser.getSelectedFile());
        }
      }
    });
    return b;
  }
}
