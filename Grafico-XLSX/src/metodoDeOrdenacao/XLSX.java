package metodoDeOrdenacao;

import java.awt.BorderLayout;
import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.border.TitledBorder;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



import javax.swing.border.LineBorder;
import java.awt.Color;
import javax.swing.JTextField;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JButton;
import javax.swing.JFileChooser;

import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.awt.event.ActionEvent;
import javax.swing.JComboBox;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JRadioButton;
import javax.swing.ButtonGroup;

public class XLSX extends JFrame {

	private JPanel contentPane;
	private JTextField txtFile;
	private String arquivo = "";
	private final ButtonGroup buttonGroup = new ButtonGroup();

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					XLSX frame = new XLSX();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public XLSX() {
		setTitle("M\u00E9todos de Ordena\u00E7\u00E3o para Planilhas do Excel");
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 615, 379);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);

		JPanel panel = new JPanel();
		panel.setBorder(new TitledBorder(new LineBorder(new Color(0, 0, 0), 2), "Escolha o Arquivo...", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel.setBounds(21, 11, 568, 127);
		contentPane.add(panel);
		panel.setLayout(null);

		txtFile = new JTextField();
		txtFile.setBounds(10, 53, 436, 20);
		panel.add(txtFile);
		txtFile.setColumns(10);

		JLabel lblCaminhoDoArquivo = new JLabel("Caminho do Arquivo:");
		lblCaminhoDoArquivo.setBounds(10, 37, 181, 14);
		panel.add(lblCaminhoDoArquivo);

		JPanel panel_1 = new JPanel();
		panel_1.setBorder(new TitledBorder(new LineBorder(new Color(0, 0, 0), 2), "Escolha o M\u00E9todo de Ordena\u00E7\u00E3o", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_1.setBounds(21, 160, 568, 115);
		contentPane.add(panel_1);
		panel_1.setLayout(null);

		JLabel lblResultado = new JLabel("");
		lblResultado.setBounds(134, 79, 331, 25);
		panel_1.add(lblResultado);

		JRadioButton rbBubble = new JRadioButton("BubbleSort");
		rbBubble.setEnabled(false);
		rbBubble.setSelected(true);
		buttonGroup.add(rbBubble);
		rbBubble.setBounds(40, 34, 109, 23);
		panel_1.add(rbBubble);

		JRadioButton rbQuick = new JRadioButton("QuickSort");
		rbQuick.setEnabled(false);
		buttonGroup.add(rbQuick);
		rbQuick.setBounds(40, 63, 109, 23);
		panel_1.add(rbQuick);

		JButton btnExecutar = new JButton("Executar");
		btnExecutar.setEnabled(false);

		JButton btnAbrir = new JButton("Abrir");
		btnAbrir.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				JFileChooser fileChooser = new JFileChooser(); 
				fileChooser.setDialogTitle("Selecione o arquivo do MS-Excel");
				fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);

				FileNameExtensionFilter filter = new FileNameExtensionFilter("planilhas", "xls", "xlsx");

				fileChooser.setFileFilter(filter);
				int retorno = fileChooser.showOpenDialog(null);

				if (retorno == JFileChooser.APPROVE_OPTION) {

					File file = fileChooser.getSelectedFile();

					txtFile.setText(file.getPath());

					arquivo = file.getPath();

				}

				rbBubble.setEnabled(true);
				rbQuick.setEnabled(true);
				rbBubble.isSelected();
				btnExecutar.setEnabled(true);


			}
		});
		btnAbrir.setMnemonic('a');
		btnAbrir.setBounds(456, 52, 89, 23);
		panel.add(btnAbrir);


		btnExecutar.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				FileInputStream fisPlanilha = null;
				try {
					fisPlanilha = new FileInputStream(arquivo);

					//cria um workbook = planilha toda com todas as abas
					XSSFWorkbook workbook = new XSSFWorkbook(fisPlanilha);

					//recuperamos apenas a primeira aba ou primeira planilha
					XSSFSheet sheet = workbook.getSheetAt(0);

					//recebe a quantidade de linhas da planilha
					int linhas = sheet.getLastRowNum();

					//cria o vetor que receberá os dados da planilha de acordo com o total de linhas
					int[] planilha = new int[linhas+1]; 

					//instancia a classe ROW
					Row row ;

					//loop que varre toda a extensao da planilha e insere os valores no vetor
					for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) 
					{
						row = sheet.getRow(i);
						Cell cell = row.getCell(0);
						planilha[i] = (int) cell.getNumericCellValue();

					}

					//decisão sobre o metodo de organização...
					if (rbBubble.isSelected()) {

						//instancia a BubbleSort...
						Bubble bubble = new Bubble(); 

						//objeto criado executa no vetor o método BubbleSort... 
						bubble.Bubble(planilha);
					}							

					if (rbQuick.isSelected()) {

						//instancia a classeQuickSort...
						Quick quick = new Quick();  //comentei devido a estar ativa a bubbleSort

						//objeto criado executa no vetor o método QuickSort... 
						quick.Quick(planilha, 0, planilha.length-1);

					}    


					//loop que preenche o vetor com o metodo já executado
					for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) 
					{
						row = sheet.getRow(i);
						Cell cellDestino = row.getCell(1);
						if(cellDestino == null) 
						{
							cellDestino = sheet.getRow(i).createCell(1);
							cellDestino.setCellValue(planilha[i]);
						}
					}

					//preenche (escreve) na planilha os valores no vetor
					FileOutputStream outFile = new FileOutputStream(new File(arquivo));
					workbook.write(outFile);

					//encerra a leitura e a escrita do arquivo..
					outFile.close();
					fisPlanilha.close();    

					JOptionPane.showMessageDialog(btnExecutar, "Método de ordenação exceutado com sucesso!");


				} catch (FileNotFoundException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}

			}
		});
		btnExecutar.setMnemonic('e');
		btnExecutar.setBounds(357, 45, 144, 23);
		panel_1.add(btnExecutar);

		JButton btnCancelar = new JButton("Cancelar");
		btnCancelar.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				System.exit(1);
			}
		});
		btnCancelar.setBounds(477, 298, 89, 23);
		contentPane.add(btnCancelar);
	}
}
