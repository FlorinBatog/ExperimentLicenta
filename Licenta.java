import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.List;
import java.util.concurrent.TimeUnit;
import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.sikuli.script.FindFailed;
import org.sikuli.script.Pattern;
import org.sikuli.script.Screen;

public class Licenta {
	static XWPFDocument document = new XWPFDocument();

	public static void generareParagraf(XWPFParagraph numeParagraf, XWPFRun numeRun, String text) {
		numeParagraf = document.createParagraph();
		numeRun = numeParagraf.createRun();
		numeRun.setText(text);
	}

	public static double extragereDinExcel(JTextField text, String numeSheet, int contorRanduri, int contorColoane)
			throws InvalidFormatException, IOException {
		File fisier = new File("C:\\User\\Desktop\\Date_ Suplimentare_Firma_" + text.getText() + ".xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fisier);
		XSSFSheet sheet = workbook.getSheet(numeSheet);
		return sheet.getRow(contorRanduri).getCell(contorColoane).getNumericCellValue();

	}

	public static void Indicatorii(int nrColoana, int nrColoanaExcel, JTextField textbox, JTextField year)
			throws FindFailed {
		Screen screen = new Screen();
		Pattern pattern = new Pattern("C:\\User\\Desktop\\TotalFirme\\baraTotalFirme.PNG");
		System.setProperty("webdriver.chrome.driver", "C:\\User\\Desktop\\Risco\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(7, TimeUnit.SECONDS);
		driver.manage().window().maximize();
		driver.get("https://www.totalfirme.ro/");
		screen.click(pattern);
		screen.type(textbox.getText());
		WebElement ele = driver.findElement(By.className("titlu"));
		ele.click();
		for (int i = 1; i < 9; i++)
			for (int j = 1; j < 12; j++) {

				WebElement elementTabelul1 = driver.findElement(
						By.xpath("//div[@id='indicatori-bilant']//table[1]//tbody[1]//tr[" + i + "]//td[" + j + "]"));
				WebElement elementTabelul2 = driver.findElement(
						By.xpath("//div[@id='indicatori-bilant']//table[2]//tbody[1]//tr[" + i + "]//td[" + j + "]"));
				for (int k = 1; k < 7; k++) {
					try {
						extragereDinExcel(textbox, "FoaieDeCalcul", k, nrColoanaExcel);

					} catch (InvalidFormatException | IOException e1) {
						e1.printStackTrace();
					}
				}
			}
		try {
			FileOutputStream out;
			out = new FileOutputStream(
					new File("C:\\User\\Desktop\\Analiza financiar-contabila anul " + year.getText() + ".docx"));

			XWPFParagraph titlu = document.createParagraph();
			XWPFRun afisareTitlu = titlu.createRun();
			afisareTitlu.setText(
					"Analiza financiar-contabila a firmei " + textbox.getText() + " pentru anul " + year.getText());
			afisareTitlu.setFontSize(25);
			// FONDUL DE RULMENT
			XWPFParagraph FondRulment = null;
			XWPFRun afisareFondRulment = null;
			generareParagraf(FondRulment, afisareFondRulment, "1.Fondul de Rulment(FR)");
			XWPFParagraph formulaFondRulment = null;
			XWPFRun FondRulmentRun = null;
			double CapitaluriPermanente = Double.parseDouble(driver
					.findElement(By
							.xpath("//div[@id='indicatori-bilant']//table[2]//tbody[1]//tr[7]//td[" + nrColoana + "]"))
					.getText().replace(".", ""));
			double ActiveImobilizate = Double.parseDouble(driver
					.findElement(By
							.xpath("//div[@id='indicatori-bilant']//table[2]//tbody[1]//tr[1]//td[" + nrColoana + "]"))
					.getText().replace(".", ""));
			double z = CapitaluriPermanente - ActiveImobilizate;
			generareParagraf(formulaFondRulment, FondRulmentRun,
					"FR = Capitaluri permanente (CPERM) – Active pe termen lung (ATL) sau active imobilizate=" + ""
							+ CapitaluriPermanente + "-" + ActiveImobilizate + "=" + z);
			if (z > 0) {
				XWPFParagraph ConcluzieFondRulment = null;
				XWPFRun afisareConcluzieFondRulment = null;
				generareParagraf(ConcluzieFondRulment, afisareConcluzieFondRulment,
						"Avem un FR pozitiv ceea ce apare ca o marja de securitate financiara  care permite intreprinderii sa faca fata fara dificultate diferitelor riscuri pe termen scurt.");
			} else {
				XWPFParagraph ConcluzieFondRulment2 = null;
				XWPFRun afisareConcluzieFondRulment2 = null;
				generareParagraf(ConcluzieFondRulment2, afisareConcluzieFondRulment2,
						"Fondul de Rulment este negativ ceea ce inseamna ca firma va avea dificultati in gestionarea riscurilor pe termen scurt.");
			}
			// NEVOIA DE FOND DE RULMENT
			XWPFParagraph NevoiaDeFondRulment = null;
			XWPFRun afisareNevoiaDeFondRulment = null;
			double DatoriiPeTermenScurt = extragereDinExcel(textbox, "FoaieDeCalcul", 1, nrColoanaExcel);
			generareParagraf(NevoiaDeFondRulment, afisareNevoiaDeFondRulment, "2.Nevoia de Fond de Rulment(NFR)");
			XWPFParagraph formulaNevoiaDeFondRulment = null;
			XWPFRun NevoiaDeFondRun = null;
			double ActiveCirculante = Double.parseDouble(driver
					.findElement(By
							.xpath("//div[@id='indicatori-bilant']//table[2]//tbody[1]//tr[2]//td[" + nrColoana + "]"))
					.getText().replace(".", ""));
			double Z = ActiveCirculante - DatoriiPeTermenScurt;
			generareParagraf(formulaNevoiaDeFondRulment, NevoiaDeFondRun,
					"NFR = Active Circulante (AC) – Datorii pe termen scurt (DTS)=" + ActiveCirculante + "-"
							+ DatoriiPeTermenScurt + "=" + Z);
			if (Z > 0) {
				XWPFParagraph ConcluzieNevoiaDeFondRulment = null;
				XWPFRun afisareConcluzieNevoiaDeFondRulment = null;
				generareParagraf(ConcluzieNevoiaDeFondRulment, afisareConcluzieNevoiaDeFondRulment,
						"Valoarea trezoreriei nete este pozitiva, deci exista un excedent de finantare.");
			} else {
				XWPFParagraph ConcluzieNevoiaDeFondRulment2 = null;
				XWPFRun afisareConcluzieNevoiaDeFondRulment2 = null;
				generareParagraf(ConcluzieNevoiaDeFondRulment2, afisareConcluzieNevoiaDeFondRulment2,
						"Valoarea trezoreriei nete este negativa, deci exista un deficit de finantare.");
			}
			// TREZORERIA NETA
			XWPFParagraph TrezorerieNeta = null;
			XWPFRun afisareTrezorerieNeta = null;
			generareParagraf(TrezorerieNeta, afisareTrezorerieNeta, "3.Trezorerie Neta(TN)");
			XWPFParagraph formulaTrezorerieNeta = null;
			XWPFRun TrezorerieNetaRun = null;
			double C = z - Z;
			generareParagraf(formulaTrezorerieNeta, TrezorerieNetaRun,
					"TN = Fondul de Rulment (FR) – Nevoia de Fond de Rulment (NFR)=" + z + "-" + Z + "=" + C);
			if (C > 0) {
				XWPFParagraph ConcluzieTrezorerieNeta = null;
				XWPFRun afisareConcluzieTrezorerieNeta = null;
				generareParagraf(ConcluzieTrezorerieNeta, afisareConcluzieTrezorerieNeta,
						"Valoarea trezoreriei nete este pozitiva, deci exista un excedent de finantare.");
			} else {
				XWPFParagraph ConcluzieTrezorerieNeta2 = null;
				XWPFRun afisareConcluzieTrezorerieNeta2 = null;
				generareParagraf(ConcluzieTrezorerieNeta2, afisareConcluzieTrezorerieNeta2,
						"Valoarea trezoreriei nete este negativa, deci exista un deficit de finantare.");
			}
			// LICHIDITATE CURENTA
			XWPFParagraph LichiditateaCurenta = null;
			XWPFRun afisareLichiditateaCurenta = null;
			generareParagraf(LichiditateaCurenta, afisareLichiditateaCurenta, "4.Lichiditatea curenta (RLG)");
			XWPFParagraph formulaLichiditateaCurenta = null;
			XWPFRun LichiditateaCurentaRun = null;
			double ActiveCurente = extragereDinExcel(textbox, "FoaieDeCalcul", 6, nrColoanaExcel);
			double DatoriiCurente = extragereDinExcel(textbox, "FoaieDeCalcul", 5, nrColoanaExcel);
			double c = ActiveCurente / DatoriiCurente;
			generareParagraf(formulaLichiditateaCurenta, LichiditateaCurentaRun,
					"Lichiditatea curenta (RLG) = Active Curente ( sau active pe termen scurt ATS) / Datorii Curente (sau datorii pe termen scurt DTS)="
							+ ActiveCurente + "/" + DatoriiCurente + "=" + c);
			if (c > 0) {
				XWPFParagraph ConcluzieLichiditateaCurenta = null;
				XWPFRun afisareConcluzieLichiditateaCurenta = null;
				generareParagraf(ConcluzieLichiditateaCurenta, afisareConcluzieLichiditateaCurenta,
						"Valoarea lichiditatii curente este supraunitara, deci "
								+ "întreprinderea are capacitatea de a-şi achita datoriile exigibile pe termen scurt din activele pe termen scurt de care dispune. ");
			} else {
				XWPFParagraph ConcluzieLichiditateaCurenta2 = null;
				XWPFRun afisareConcluzieLichiditateaCurenta2 = null;
				generareParagraf(ConcluzieLichiditateaCurenta2, afisareConcluzieLichiditateaCurenta2,
						"Valoarea lichiditatii curente este subunitara, deci "
								+ "întreprinderea nu are capacitatea de a-şi achita datoriile exigibile pe termen scurt din activele pe termen scurt de care dispune. ");
			}
			// LICHIDITATE PARTIALA
			XWPFParagraph LichiditateaPartiala = null;
			XWPFRun afisareLichiditateaPartiala = null;
			generareParagraf(LichiditateaPartiala, afisareLichiditateaPartiala, "4. Lichiditatea partiala (RLP)");
			XWPFParagraph formulaLichiditateaPartiala = null;
			XWPFRun LichiditateaPartialaRun = null;
			double Stocuri = Double.parseDouble(driver
					.findElement(By
							.xpath("//div[@id='indicatori-bilant']//table[2]//tbody[1]//tr[3]//td[" + nrColoana + "]"))
					.getText().replace(".", ""));
			double diferenta = ActiveCurente - Stocuri;
			double F = diferenta / DatoriiCurente;
			generareParagraf(formulaLichiditateaPartiala, LichiditateaPartialaRun,
					"RLP = (Active Curente – Stocuri) / Datorii Curente=(" + ActiveCurente + "-" + Stocuri + ")/"
							+ DatoriiCurente + "=" + F);
			if (F > 0) {
				XWPFParagraph ConcluzieLichiditateaPartiala = null;
				XWPFRun afisareConcluzieLichiditateaPartiala = null;
				generareParagraf(ConcluzieLichiditateaPartiala, afisareConcluzieLichiditateaPartiala,
						"Valoarea trezoreriei nete este pozitiva, deci exista un excedent de finantare.");
			} else {
				XWPFParagraph ConcluzieLichiditateaPartiala2 = null;
				XWPFRun afisareConcluzieLichiditateaPartiala2 = null;
				generareParagraf(ConcluzieLichiditateaPartiala2, afisareConcluzieLichiditateaPartiala2,
						"Valoarea trezoreriei nete este negativa, deci exista un deficit de finantare.");
			}
			// LICHIDITATE IMEDIATA
			XWPFParagraph LichiditateaImediata = null;
			XWPFRun afisareLichiditateaImediata = null;
			generareParagraf(LichiditateaImediata, afisareLichiditateaImediata, "5. Lichiditatea imediata (RLI)");
			XWPFParagraph formulaLichiditateaImediata = null;
			XWPFRun LichiditateaImediataRun = null;
			double Disponibilitati = Double.parseDouble(
					driver.findElement(By.xpath("//div[@id='indicatori-bilant']//table[1]//tbody[1]//tr[6]//td[2]"))
							.getText().replace(".", ""));
			double f = Disponibilitati / DatoriiCurente;
			generareParagraf(formulaLichiditateaImediata, LichiditateaImediataRun,
					"RLI= Disponibilitati/ Datorii curente (DC) sau datorii pe termen scurt (DTS)=" + Disponibilitati
							+ "/" + DatoriiCurente + "=" + f);
			if (f > 0) {
				XWPFParagraph ConcluzieLichiditateaImediata = null;
				XWPFRun afisareConcluzieLichiditateaImediata = null;
				generareParagraf(ConcluzieLichiditateaImediata, afisareConcluzieLichiditateaImediata,
						"Societatea poate sa-si ramburseze datoriile pe termen scurt din disponibilul existent");
			} else {
				XWPFParagraph ConcluzieLichiditateaImediata2 = null;
				XWPFRun afisareConcluzieLichiditateaImediata2 = null;
				generareParagraf(ConcluzieLichiditateaImediata2, afisareConcluzieLichiditateaImediata2,
						"Societatea nu poate sa-si ramburseze datoriile pe termen scurt din disponibilul existent");
			}
			// RATA SOLVABILITATII
			XWPFParagraph RataSolvabilitatii = null;
			XWPFRun afisareRataSolvabilitatii = null;
			generareParagraf(RataSolvabilitatii, afisareRataSolvabilitatii, "6. Rata Solvabilitatii (RS)");
			XWPFParagraph formulaRataSolvabilitatii = null;
			XWPFRun RataSolvabilitatiiRun = null;
			double ActivTotal = ActiveImobilizate + ActiveCirculante;
			double DatoriiTotale = Double.parseDouble(driver
					.findElement(By
							.xpath("//div[@id='indicatori-bilant']//table[1]//tbody[1]//tr[8]//td[" + nrColoana + "]"))
					.getText().replace(".", ""));
			double I = ActivTotal / DatoriiTotale;
			generareParagraf(formulaRataSolvabilitatii, RataSolvabilitatiiRun,
					"RS= Activ total/ Datorii totale= " + ActivTotal + "/" + DatoriiTotale + "=" + I);
			if (I > 0) {
				XWPFParagraph ConcluzieRataSolvabilitatii = null;
				XWPFRun afisareConcluzieRataSolvabilitatii = null;
				generareParagraf(ConcluzieRataSolvabilitatii, afisareConcluzieRataSolvabilitatii,
						"Valoarea ratei este mai mare de 1,5 "
								+ "fapt ce evidentiaza că întreprinderea este solvabilă, adică are capacitatea de a-şi achita datoriile pe termen scurt, mediu şi lung prin valorificarea activelor de care dispune.");
			} else {
				XWPFParagraph ConcluzieRataSolvabilitatii2 = null;
				XWPFRun afisareConcluzieRataSolvabilitatii2 = null;
				generareParagraf(ConcluzieRataSolvabilitatii2, afisareConcluzieRataSolvabilitatii2,
						"Valoarea ratei este mai mica de 1,5"
								+ "fapt ce evidentiaza că întreprinderea nu are capacitatea de a-şi achita datoriile pe termen scurt, mediu şi lung prin valorificarea activelor de care dispune.");
			}
			// COEFICIENTUL TOTAL DE INDATORARE
			XWPFParagraph CoeficientulTotalDeIndatorare = null;
			XWPFRun afisareCoeficientulTotalDeIndatorare = null;
			generareParagraf(CoeficientulTotalDeIndatorare, afisareCoeficientulTotalDeIndatorare,
					"7. Coeficientul total de indatorare (CTÎ)");
			XWPFParagraph formulaCoeficientulTotalDeIndatorare = null;
			XWPFRun CoeficientulTotalDeIndatorareRun = null;
			double Capitaluri = Double.parseDouble(driver
					.findElement(By
							.xpath("//div[@id='indicatori-bilant']//table[2]//tbody[1]//tr[8]//td[" + nrColoana + "]"))
					.getText().replace(",", ""));
			double Rezerve = extragereDinExcel(textbox, "FoaieDeCalcul", 3, nrColoanaExcel);
			double RezultatReportat = extragereDinExcel(textbox, "FoaieDeCalcul", 4, nrColoanaExcel);
			double CapitaluriProprii = Capitaluri + Rezerve + RezultatReportat;
			double i1 = DatoriiTotale / CapitaluriProprii;
			generareParagraf(formulaCoeficientulTotalDeIndatorare, CoeficientulTotalDeIndatorareRun,
					"CTÎ = Datorii totale/ Capitaluri proprii=" + DatoriiTotale + "/" + CapitaluriProprii + "=" + i1);
			if (i1 > 0) {
				XWPFParagraph ConcluzieCoeficientulTotalDeIndatorare = null;
				XWPFRun afisareConcluzieCoeficientulTotalDeIndatorare = null;
				generareParagraf(ConcluzieCoeficientulTotalDeIndatorare, afisareConcluzieCoeficientulTotalDeIndatorare,
						"Societatea poate contracta împrumuturi noi, deoarece valoarea coeficientului de îndatorare este mai mic decat 2. ");
			} else {
				XWPFParagraph ConcluzieLichiditateaPartiala2 = null;
				XWPFRun afisareConcluzieLichiditateaPartiala2 = null;
				generareParagraf(ConcluzieLichiditateaPartiala2, afisareConcluzieLichiditateaPartiala2,
						"Societatea nu poate contracta împrumuturi noi, deoarece valoarea coeficientului de îndatorare este mai mare decat 2. ");
			}
			document.write(out);
			out.close();
			System.out.println("Documentul a fost creat cu succes");
			driver.close();

		} catch (IOException e1) {
			e1.printStackTrace();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		}
	}

	public Licenta() throws FindFailed, InterruptedException, IOException {

		JFrame frame = new JFrame();
		JPanel panel = new JPanel();
		JButton button = new JButton("BUTON");
		JLabel label = new JLabel("Buna ziua! Va rugam sa introduceti numele firmei pe care doriti sa o analizam");
		JTextField textbox = new JTextField("Nume firma");
		JTextField year = new JTextField("Anul");
		panel.setBorder(BorderFactory.createEmptyBorder(10, 20, 10, 20));
		panel.add(label, BorderLayout.CENTER);
		panel.add(textbox);
		panel.add(year);
		panel.add(button, BorderLayout.CENTER);
		frame.setMinimumSize(new Dimension(500, 150));
		frame.add(panel, BorderLayout.CENTER);
		frame.setTitle("Test GUI");
		frame.setVisible(true);
		button.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				Thread thread = new Thread(new Runnable() {

					@Override
					public void run() {
						try {

							switch (year.getText()) {
							case "2019":
								Indicatorii(2, 1, textbox, year);
								break;

							case "2018":

								Indicatorii(3, 2, textbox, year);
								break;

							case "2017":

								Indicatorii(4, 3, textbox, year);
								break;

							case "2016":

								Indicatorii(5, 4, textbox, year);
								break;

							case "2015":

								Indicatorii(6, 5, textbox, year);
								break;

							case "2014":
								Indicatorii(7, 6, textbox, year);
								break;

							case "2013":

								Indicatorii(8, 7, textbox, year);
								break;

							case "2012":

								Indicatorii(9, 8, textbox, year);
								break;

							case "2011":

								Indicatorii(10, 9, textbox, year);
								break;

							case "2010":

								Indicatorii(11, 10, textbox, year);
								break;
							}
						} catch (FindFailed e2) {
							e2.printStackTrace();
						}
					}

				});
				thread.start();

			}
		});

	}

	public static void main(String[] args) throws FindFailed, InterruptedException, IOException {
		new Licenta();
	}

}
