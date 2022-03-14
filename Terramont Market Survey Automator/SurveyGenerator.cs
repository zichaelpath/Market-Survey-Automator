using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using AWord = Aspose.Words;
using Aspose.Words.Drawing;

namespace Terramont_Market_Survey_Automator
{
	public class SurveyGenerator
	{
		private AWord.Document aDoc;
		private AWord.DocumentBuilder aBuilder;
		private List<Property> surveyProperties;
		private List<string> surveyFloorPlans;
		private List<string> surveyGeneralImages;
		private List<string> clientNeeds;
		private List<string> imageFilePaths;
		private bool frenchMode;
		private bool rem;
		public SurveyGenerator(List<Property> properties, List<string> floorPlans, List<string> generalImages,
			List<string> needs, List<string> images, bool isFrench, bool hasRem)
		{
			aDoc = new AWord.Document();
			aBuilder = new AWord.DocumentBuilder(aDoc);
			
			surveyProperties = properties;
			surveyFloorPlans = floorPlans;
			surveyGeneralImages = generalImages;
			clientNeeds = needs;
			imageFilePaths = images;
			frenchMode = isFrench;
			rem = hasRem;
		}
		public void UpdateFooter()
		{
			AWord.ParagraphFormat format = aBuilder.ParagraphFormat;
			aBuilder.MoveToHeaderFooter(AWord.HeaderFooterType.FooterPrimary);
			format = aBuilder.ParagraphFormat;
			format.Shading.BackgroundPatternColor = Color.DarkBlue;
			format.Style.Font.Color = Color.White;
			format.Style.Font.Size = 16;
			aBuilder.InsertField("PAGE");
			aBuilder.Write("                          Services Immobiliers Terramont Real Estate Services");
		}
		public void CreateSurvey()
		{
			GenerateNeedsAnalysisPage();
			GenerateResearchProcessText();
			if (rem)
			{
				GenerateRemPage();
			}
			GenerateAvailabilityOvervew();
			GenerateSiteDetailsTables();
			UpdateFooter();
			string directory = @"C:\Terramont Clients\" + clientNeeds[8] + "\\Market Survey.docx";

			aDoc.Save(directory);
		}
		public void GenerateRemPage()
		{
			aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
			aBuilder.InsertImage(@"C:\Market Survey Images\rem.png");
			aBuilder.InsertBreak(AWord.BreakType.LineBreak);

			
			
			aBuilder.InsertImage(@"C:\Market Survey Images\remMap.png");
			
			aBuilder.Font.Size = 9;
			aBuilder.Writeln("The new REM is an integrated network linking downtown Montreal, South Shore, West Island, North Shore and the airport.  It is estimated to be completed in 2023, will have 26 stations, will be 67 kilometers, and will run 20 hours a day 7 days a week.\n\n");
			aBuilder.Writeln("Once completed, the REM will be one of the largest automated transportation systems in the world after Singapore, Dubai and Vancouver. As a single, integrated transportation network, the REM will offer a number of efficient travel options in the Greater Montréal area. It will be connected with bus networks, commuter trains (Mascouche and Saint-Hilaire lines) and with the Montréal metro (Blue, Green and Orange lines).");
			aBuilder.Font.Bold = true;
			aBuilder.Writeln("Potential Economic Benefits");
			aBuilder.Font.Bold = false;
			aBuilder.Writeln("•	GDP: the REM could potentially add more than $3.7 billion to Québec’s GDP over four years\n•	Real estate developments: close to $5 billion in private real estate developments along the route are currently expected\n•	Jobs: more than 34,000 direct and indirect jobs will be created during the construction phase and more than 1,000 permanent jobs will be created once the REM starts running\n•	Environment: the REM could help reduce GHG emissions by  680,000 tons over 25 years of operation and accelerate Québec’s transition to a low-carbon economy\n•	Traffic congestion: this new public transit system could reduce economic losses associated with traffic congestion, currently estimated at $1.9 billion annually in the Greater Montréal area");
			aBuilder.Font.Bold = true;
			aBuilder.Writeln("Sustainable Mobility Services");
			aBuilder.Font.Bold = false;
			aBuilder.Writeln("Agreements have been reached with various sustainable mobility services (Bixi, Car2go, Communauto, Netlift, Téo Taxi, Vélo Québec) to offer passengers a variety of travel options between home and the station. For additional options, mobility services around REM stations will combine car-sharing, carpooling and electric taxi services.");
			

			aBuilder.InsertBreak(AWord.BreakType.PageBreak);
		}
		public void GenerateNeedsAnalysisPage()
		{
			

			AWord.Font font = aBuilder.Font;
			font.Size = 24;
			font.Bold = false;

			if (!frenchMode)
			{
				AWord.ParagraphFormat format = aBuilder.ParagraphFormat;
				format.Alignment = AWord.ParagraphAlignment.Left;

				aBuilder.InsertImage(@"C:\Market Survey Images\needsAnalysis.png");
				aBuilder.InsertBreak(AWord.BreakType.LineBreak);
				aBuilder.Write("\n");
				format.Alignment = AWord.ParagraphAlignment.Right;

				aBuilder.ParagraphFormat.StyleIdentifier = AWord.StyleIdentifier.Heading1;
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
				aBuilder.Font.Color = Color.LightBlue;
				aBuilder.Writeln("Needs\nAnalysis");
				aBuilder.Font.Color = Color.Black;
				aBuilder.ParagraphFormat.StyleIdentifier = AWord.StyleIdentifier.Normal;
				font.Size = 16;


				AWord.Tables.Table table = aBuilder.StartTable();
				format.Alignment = AWord.ParagraphAlignment.Left;
				aBuilder.CellFormat.Borders.LineStyle = AWord.LineStyle.Single;
				aBuilder.CellFormat.TopPadding = 10;
				aBuilder.CellFormat.BottomPadding = 10;
				aBuilder.InsertCell();
				aBuilder.CellFormat.HorizontalMerge = AWord.Tables.CellMerge.First;
				table.PreferredWidth = AWord.Tables.PreferredWidth.FromPercent(100);
				aBuilder.Write("Requirements");
				aBuilder.InsertCell();
				aBuilder.CellFormat.HorizontalMerge = AWord.Tables.CellMerge.Previous;
				aBuilder.EndRow();


				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
				aBuilder.Write("Area ft\xB2");
				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
				aBuilder.Write(clientNeeds[0]);
				aBuilder.EndRow();

				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
				aBuilder.Write("Term");
				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
				aBuilder.Write(clientNeeds[1]);

				aBuilder.EndRow();

				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
				aBuilder.Write("Growth");
				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
				aBuilder.Write(clientNeeds[2]);

				aBuilder.EndRow();

				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
				aBuilder.Write("Relocation Objectives");
				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
				aBuilder.Write(clientNeeds[3]);

				aBuilder.EndRow();

				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
				aBuilder.Write("Location");
				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
				aBuilder.Write(clientNeeds[4]);

				aBuilder.EndRow();

				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
				aBuilder.Write("Building Type");
				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
				aBuilder.Write(clientNeeds[5]);

				aBuilder.EndRow();

				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
				aBuilder.Write("Parking");
				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
				aBuilder.Write(clientNeeds[6]);

				aBuilder.EndRow();

				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
				aBuilder.Write("Comments");
				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
				aBuilder.Write(clientNeeds[7]);

				aBuilder.EndRow();
				aBuilder.EndTable();
				//UpdateFooter();
				aBuilder.InsertBreak(AWord.BreakType.PageBreak);
			}
			else
			{
				AWord.ParagraphFormat format = aBuilder.ParagraphFormat;
				
				format.Alignment = AWord.ParagraphAlignment.Right;

				aBuilder.ParagraphFormat.StyleIdentifier = AWord.StyleIdentifier.Heading1;
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
				aBuilder.Font.Color = Color.LightBlue;
				aBuilder.Writeln("Analyse des besoins");
				aBuilder.Font.Color = Color.Black;
				aBuilder.ParagraphFormat.StyleIdentifier = AWord.StyleIdentifier.Normal;
				font.Size = 16;


				AWord.Tables.Table table = aBuilder.StartTable();
				format.Alignment = AWord.ParagraphAlignment.Left;
				aBuilder.CellFormat.Borders.LineStyle = AWord.LineStyle.Single;
				aBuilder.CellFormat.TopPadding = 10;
				aBuilder.CellFormat.BottomPadding = 10;
				aBuilder.InsertCell();
				aBuilder.CellFormat.HorizontalMerge = AWord.Tables.CellMerge.First;
				table.PreferredWidth = AWord.Tables.PreferredWidth.FromPercent(100);
				aBuilder.Write("Exigences");
				aBuilder.InsertCell();
				aBuilder.CellFormat.HorizontalMerge = AWord.Tables.CellMerge.Previous;
				aBuilder.EndRow();


				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
				aBuilder.Write("Superficie");
				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
				aBuilder.Write(clientNeeds[0]);
				aBuilder.EndRow();

				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
				aBuilder.Write("Terme");
				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
				aBuilder.Write(clientNeeds[1]);

				aBuilder.EndRow();

				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
				aBuilder.Write("Croissance");
				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
				aBuilder.Write(clientNeeds[2]);

				aBuilder.EndRow();

				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
				aBuilder.Write("Objectifs");
				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
				aBuilder.Write(clientNeeds[3]);

				aBuilder.EndRow();

				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
				aBuilder.Write("Emplacement");
				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
				aBuilder.Write(clientNeeds[4]);

				aBuilder.EndRow();

				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
				aBuilder.Write("Type de Bâtiment");
				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
				aBuilder.Write(clientNeeds[5]);

				aBuilder.EndRow();

				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
				aBuilder.Write("Parking");
				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
				aBuilder.Write(clientNeeds[6]);

				aBuilder.EndRow();

				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
				aBuilder.Write("Critères de notre recherche");
				aBuilder.InsertCell();
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
				aBuilder.Write(clientNeeds[7]);

				aBuilder.EndRow();
				aBuilder.EndTable();
				//UpdateFooter();
				aBuilder.InsertBreak(AWord.BreakType.PageBreak);
			}
			

		}
		public void GenerateResearchProcessText()
		{
			if (!frenchMode)
			{
				string firstParagraph = "Following our meeting we completed a real estate analysis amongst all the availabilities in the market, our process is described as follows:";
				string bulletPointOne = "- Thorough analysis of " + clientNeeds[8] + " Requirements, including discussions on current location, as it relates to improvements";
				bulletPointOne += ", financial and space goals, etc.\n";

				string bulletPointTwo = "- Greater Montreal search of all availabilities (websites, databases and mass emailing to all brokers and landlords to ensure no availabilities are missed).\n";
				string bulletPointThree = "- Discussions with all major landlords and brokers with potential availabilities.\n";
				string bulletPointFour = "- Prequalifying the spaces with targeted tours of sites with greater potential to fulfill requirements.\n";
				string bulletPointFive = "- Analysis of all NAI Terramont Commercial and all other broker availabilities.";



				string lastParagraph = "Our property market search is based on satisfying all of " + clientNeeds[8] + "'s corporate goals in order to provide the best working environment, " +
										"both physically and geographically and to help you and your team reach its maximum productivity and employee retention.";


				AWord.Font font = aBuilder.Font;
				font.Size = 24;
				font.Bold = false;

				AWord.ParagraphFormat format = aBuilder.ParagraphFormat;
				format.Alignment = AWord.ParagraphAlignment.Left;

				aBuilder.InsertImage(@"C:\Market Survey Images\researchProcess.png");
				aBuilder.InsertBreak(AWord.BreakType.LineBreak);
				aBuilder.Writeln("");
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
				aBuilder.Font.Color = Color.LightBlue;
				aBuilder.ParagraphFormat.StyleIdentifier = AWord.StyleIdentifier.Heading2;
				aBuilder.Writeln("Research\nProcess");
				aBuilder.Font.Color = Color.Black;
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
				aBuilder.ParagraphFormat.StyleIdentifier = AWord.StyleIdentifier.Normal;
				font.Size = 16;
				font.Bold = true;
				aBuilder.Writeln(firstParagraph);

				font.Bold = false;
				aBuilder.Writeln(bulletPointOne + bulletPointTwo + bulletPointThree + bulletPointFour + bulletPointFive);

				font.Bold = true;

				aBuilder.Writeln(lastParagraph);

				//UpdateFooter();
				aBuilder.InsertBreak(AWord.BreakType.PageBreak);
				aBuilder.InsertImage(@"C:\Market Survey Images\researchProcess.png");
				aBuilder.InsertBreak(AWord.BreakType.LineBreak);
				aBuilder.InsertBreak(AWord.BreakType.PageBreak);
			}
			else
			{
				string firstParagraph = "Suite à notre rencontre nous avons réalisé une analyse immobilière parmi toutes les disponibilités du marché, notre démarche se décrit comme suit :\n";
				string bulletPointOne = "- Analyse approfondie des exigences de " + clientNeeds[8] + ", y compris des discussions sur l'emplacement actuel, en ce qui concerne les améliorations";
				bulletPointOne += ", les objectifs financiers et d'espace, etc.\n";

				string bulletPointTwo = "- Recherche de toutes les disponibilités dans " + clientNeeds[4] + " (sites Web, bases de données et envoi massif de courriels à tous les courtiers et propriétaires pour s'assurer qu'aucune disponibilité ne manque). \n";
				string bulletPointThree = "- Discussions avec tous les principaux propriétaires et courtiers avec des disponibilités potentielles.\n";
				string bulletPointFour = "- Prequalifying the spaces with targeted tours of sites with greater potential to fulfill requirements.\n";
				string bulletPointFive = "- Préqualifier les espaces avec des visites ciblées de sites à plus fort potentiel de satisfaction des besoins.\n";



				string lastParagraph = "Notre recherche de marché immobilier est basée sur la satisfaction de tous les objectifs d'entreprise de " + clientNeeds[8] +" afin de fournir le meilleur environnement de travail, à la fois physiquement et géographiquement et pour vous aider, vous et votre équipe, à atteindre sa productivité maximale et la rétention des employés. ";


				AWord.Font font = aBuilder.Font;
				font.Size = 24;
				font.Bold = false;

				AWord.ParagraphFormat format = aBuilder.ParagraphFormat;
				format.Alignment = AWord.ParagraphAlignment.Left;

				
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
				aBuilder.Font.Color = Color.LightBlue;
				aBuilder.ParagraphFormat.StyleIdentifier = AWord.StyleIdentifier.Heading2;
				aBuilder.Writeln("Processus de recherche");
				aBuilder.Font.Color = Color.Black;
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
				aBuilder.ParagraphFormat.StyleIdentifier = AWord.StyleIdentifier.Normal;
				font.Size = 16;
				font.Bold = true;
				aBuilder.Writeln(firstParagraph);
				aBuilder.Write("\n");
				font.Bold = false;
				aBuilder.Writeln(bulletPointOne + bulletPointTwo + bulletPointThree + bulletPointFour + bulletPointFive);
				aBuilder.Write("\n");
				font.Bold = true;

				aBuilder.Writeln(lastParagraph);

				//UpdateFooter();
				aBuilder.InsertBreak(AWord.BreakType.PageBreak);
				aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
				aBuilder.Font.Color = Color.LightBlue;
				aBuilder.ParagraphFormat.StyleIdentifier = AWord.StyleIdentifier.Heading2;
				aBuilder.Writeln("Processus de recherche");
				aBuilder.InsertBreak(AWord.BreakType.PageBreak);
			}
			
		}
		public void GenerateAvailabilityOvervew()
		{
			AWord.Font font = aBuilder.Font;
			font.Size = 24;
			font.Bold = false;

			if (!frenchMode)
			{
				AWord.ParagraphFormat format = aBuilder.ParagraphFormat;
				format.Alignment = AWord.ParagraphAlignment.Left;

				aBuilder.InsertImage(@"C:\Market Survey Images\availabilityOverview.png");
				aBuilder.InsertBreak(AWord.BreakType.LineBreak);

				font.Size = 11;
				font.Color = Color.White;
				AWord.Tables.Table firstHalfTable = aBuilder.StartTable();


				aBuilder.InsertCell();
				aBuilder.Write(" ");
				firstHalfTable.Rows[0].Cells[0].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
				aBuilder.InsertCell();
				aBuilder.Write("Address");
				firstHalfTable.Rows[0].Cells[1].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
				aBuilder.InsertCell();
				aBuilder.Write("Administrator Landlord");
				firstHalfTable.Rows[0].Cells[2].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
				aBuilder.InsertCell();
				aBuilder.Write("Rentable Area");
				firstHalfTable.Rows[0].Cells[3].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
				aBuilder.InsertCell();
				aBuilder.Write("Term");
				firstHalfTable.Rows[0].Cells[4].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
				aBuilder.InsertCell();
				aBuilder.Write("Occupancy");
				firstHalfTable.Rows[0].Cells[5].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
				aBuilder.InsertCell();
				aBuilder.Write("Incentives");
				firstHalfTable.Rows[0].Cells[6].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
				aBuilder.InsertCell();
				aBuilder.Write("Net Rent");
				firstHalfTable.Rows[0].Cells[7].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
				aBuilder.EndRow();



				int propCount = 1;
				foreach (Property property in surveyProperties)
				{
					int RowCount = propCount;
					font.Color = Color.White;
					aBuilder.InsertCell();
					aBuilder.Write(propCount.ToString());
					firstHalfTable.Rows[RowCount].Cells[0].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;


					font.Color = Color.Black;
					aBuilder.InsertCell();
					aBuilder.Write(property.Address);

					if (propCount % 2 == 0)
					{
						firstHalfTable.Rows[RowCount].Cells[1].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
					}
					else
					{
						firstHalfTable.Rows[RowCount].Cells[1].CellFormat.Shading.BackgroundPatternColor = Color.White;
					}

					aBuilder.InsertCell();
					aBuilder.Write(property.Landlord);

					if (propCount % 2 == 0)
					{
						firstHalfTable.Rows[RowCount].Cells[2].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
					}
					else
					{
						firstHalfTable.Rows[RowCount].Cells[2].CellFormat.Shading.BackgroundPatternColor = Color.White;
					}

					aBuilder.InsertCell();
					aBuilder.Write(property.RentableArea);

					if (propCount % 2 == 0)
					{
						firstHalfTable.Rows[RowCount].Cells[3].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
					}
					else
					{
						firstHalfTable.Rows[RowCount].Cells[3].CellFormat.Shading.BackgroundPatternColor = Color.White;
					}

					aBuilder.InsertCell();
					aBuilder.Write(property.Term);

					if (propCount % 2 == 0)
					{
						firstHalfTable.Rows[RowCount].Cells[4].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
					}
					else
					{
						firstHalfTable.Rows[RowCount].Cells[4].CellFormat.Shading.BackgroundPatternColor = Color.White;
					}

					aBuilder.InsertCell();
					aBuilder.Write(property.Occupancy);

					if (propCount % 2 == 0)
					{
						firstHalfTable.Rows[RowCount].Cells[5].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
					}
					else
					{
						firstHalfTable.Rows[RowCount].Cells[5].CellFormat.Shading.BackgroundPatternColor = Color.White;
					}

					aBuilder.InsertCell();
					aBuilder.Write(property.Incentives);

					if (propCount % 2 == 0)
					{
						firstHalfTable.Rows[RowCount].Cells[6].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
					}
					else
					{
						firstHalfTable.Rows[RowCount].Cells[6].CellFormat.Shading.BackgroundPatternColor = Color.White;
					}

					aBuilder.InsertCell();
					aBuilder.Write(property.NetRent);

					if (propCount % 2 == 0)
					{
						firstHalfTable.Rows[RowCount].Cells[7].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
					}
					else
					{
						firstHalfTable.Rows[RowCount].Cells[7].CellFormat.Shading.BackgroundPatternColor = Color.White;
					}

					aBuilder.EndRow();
					propCount++;
				}

				aBuilder.EndTable();
				//UpdateFooter();
				aBuilder.InsertBreak(AWord.BreakType.PageBreak);


				font.Color = Color.White;
				AWord.Tables.Table secondHalfTable = aBuilder.StartTable();
				aBuilder.InsertCell();
				aBuilder.Write(" ");
				secondHalfTable.Rows[0].Cells[0].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
				aBuilder.InsertCell();
				aBuilder.Write("Operating Costs");
				secondHalfTable.Rows[0].Cells[1].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
				aBuilder.InsertCell();
				aBuilder.Write("Taxes");
				secondHalfTable.Rows[0].Cells[2].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
				aBuilder.InsertCell();
				aBuilder.Write("Energy");
				secondHalfTable.Rows[0].Cells[3].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
				aBuilder.InsertCell();
				aBuilder.Write("Total Additional Rent");
				secondHalfTable.Rows[0].Cells[4].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
				aBuilder.InsertCell();
				aBuilder.Write("Gross Rent");
				secondHalfTable.Rows[0].Cells[5].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
				aBuilder.InsertCell();
				aBuilder.Write("Parking");
				secondHalfTable.Rows[0].Cells[6].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
				aBuilder.InsertCell();
				aBuilder.Write("Comments");
				secondHalfTable.Rows[0].Cells[7].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;

				aBuilder.EndRow();


				propCount = 1;
				foreach (Property property in surveyProperties)
				{
					int RowCount = propCount;
					font.Color = Color.White;
					aBuilder.InsertCell();
					aBuilder.Write(propCount.ToString());
					secondHalfTable.Rows[RowCount].Cells[0].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
					aBuilder.Write(propCount.ToString());

					font.Color = Color.Black;
					aBuilder.InsertCell();
					aBuilder.Write(property.OperationCosts);

					if (propCount % 2 == 0)
					{
						secondHalfTable.Rows[RowCount].Cells[1].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
					}
					else
					{
						secondHalfTable.Rows[RowCount].Cells[1].CellFormat.Shading.BackgroundPatternColor = Color.White;
					}

					aBuilder.InsertCell();
					aBuilder.Write(property.Taxes);

					if (propCount % 2 == 0)
					{
						secondHalfTable.Rows[RowCount].Cells[2].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
					}
					else
					{
						secondHalfTable.Rows[RowCount].Cells[2].CellFormat.Shading.BackgroundPatternColor = Color.White;
					}

					aBuilder.InsertCell();
					aBuilder.Write(property.Energy);

					if (propCount % 2 == 0)
					{
						secondHalfTable.Rows[RowCount].Cells[3].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
					}
					else
					{
						secondHalfTable.Rows[RowCount].Cells[3].CellFormat.Shading.BackgroundPatternColor = Color.White;
					}

					aBuilder.InsertCell();
					aBuilder.Write(property.TotalAdditionalRent);

					if (propCount % 2 == 0)
					{
						secondHalfTable.Rows[RowCount].Cells[4].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
					}
					else
					{
						secondHalfTable.Rows[RowCount].Cells[4].CellFormat.Shading.BackgroundPatternColor = Color.White;
					}

					aBuilder.InsertCell();
					aBuilder.Write(property.GrossRent);

					if (propCount % 2 == 0)
					{
						secondHalfTable.Rows[RowCount].Cells[5].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
					}
					else
					{
						secondHalfTable.Rows[RowCount].Cells[5].CellFormat.Shading.BackgroundPatternColor = Color.White;
					}

					aBuilder.InsertCell();
					aBuilder.Write(property.Parking);

					if (propCount % 2 == 0)
					{
						secondHalfTable.Rows[RowCount].Cells[6].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
					}
					else
					{
						secondHalfTable.Rows[RowCount].Cells[6].CellFormat.Shading.BackgroundPatternColor = Color.White;
					}

					aBuilder.InsertCell();
					aBuilder.Write(property.Comments);

					if (propCount % 2 == 0)
					{
						secondHalfTable.Rows[RowCount].Cells[7].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
					}
					else
					{
						secondHalfTable.Rows[RowCount].Cells[7].CellFormat.Shading.BackgroundPatternColor = Color.White;
					}

					aBuilder.EndRow();
					propCount++;
				}

				aBuilder.EndTable();
			}
			else
			{
				AWord.ParagraphFormat format = aBuilder.ParagraphFormat;
				format.Alignment = AWord.ParagraphAlignment.Left;
				aBuilder.Font.Size = 24;
				aBuilder.Writeln("Sommaire des disponibilités");

				font.Size = 11;
				font.Color = Color.White;
				AWord.Tables.Table firstHalfTable = aBuilder.StartTable();

				aBuilder.InsertCell();
				aBuilder.Write(" ");
				firstHalfTable.Rows[0].Cells[0].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
				aBuilder.InsertCell();
				aBuilder.Write("Adresse");
			firstHalfTable.Rows[0].Cells[1].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
			aBuilder.InsertCell();
			aBuilder.Write("Administrateur");
			firstHalfTable.Rows[0].Cells[2].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
			aBuilder.InsertCell();
			aBuilder.Write("Superficie (pi²)");
			firstHalfTable.Rows[0].Cells[3].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
			aBuilder.InsertCell();
			aBuilder.Write("Terme du bail");
			firstHalfTable.Rows[0].Cells[4].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
			aBuilder.InsertCell();
			aBuilder.Write("Occupation");
			firstHalfTable.Rows[0].Cells[5].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
			aBuilder.InsertCell();
			aBuilder.Write("Allocation ($ par pi²)");
			firstHalfTable.Rows[0].Cells[6].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
			aBuilder.InsertCell();
			aBuilder.Write("Loyer Net");
			firstHalfTable.Rows[0].Cells[7].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
			aBuilder.EndRow();
			
			

			int propCount = 1;
			foreach (Property property in surveyProperties)
			{
				int RowCount = propCount;
				font.Color = Color.White;
				aBuilder.InsertCell();
				aBuilder.Write(propCount.ToString());
				firstHalfTable.Rows[RowCount].Cells[0].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
				
				
				font.Color = Color.Black;
				aBuilder.InsertCell();
				aBuilder.Write(property.Address);

				if (propCount % 2 == 0)
				{
					firstHalfTable.Rows[RowCount].Cells[1].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
				}
				else
				{
					firstHalfTable.Rows[RowCount].Cells[1].CellFormat.Shading.BackgroundPatternColor = Color.White;
				}

				aBuilder.InsertCell();
				aBuilder.Write(property.Landlord);

				if (propCount % 2 == 0)
				{
					firstHalfTable.Rows[RowCount].Cells[2].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
				}
				else
				{
					firstHalfTable.Rows[RowCount].Cells[2].CellFormat.Shading.BackgroundPatternColor = Color.White;
				}

				aBuilder.InsertCell();
				aBuilder.Write(property.RentableArea);

				if (propCount % 2 == 0)
				{
					firstHalfTable.Rows[RowCount].Cells[3].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
				}
				else
				{
					firstHalfTable.Rows[RowCount].Cells[3].CellFormat.Shading.BackgroundPatternColor = Color.White;
				}

				aBuilder.InsertCell();
				aBuilder.Write(property.Term);

				if (propCount % 2 == 0)
				{
					firstHalfTable.Rows[RowCount].Cells[4].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
				}
				else
				{
					firstHalfTable.Rows[RowCount].Cells[4].CellFormat.Shading.BackgroundPatternColor = Color.White;
				}

				aBuilder.InsertCell();
				aBuilder.Write(property.Occupancy);

				if (propCount % 2 == 0)
				{
					firstHalfTable.Rows[RowCount].Cells[5].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
				}
				else
				{
					firstHalfTable.Rows[RowCount].Cells[5].CellFormat.Shading.BackgroundPatternColor = Color.White;
				}

				aBuilder.InsertCell();
				aBuilder.Write(property.Incentives);

				if (propCount % 2 == 0)
				{
					firstHalfTable.Rows[RowCount].Cells[6].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
				}
				else
				{
					firstHalfTable.Rows[RowCount].Cells[6].CellFormat.Shading.BackgroundPatternColor = Color.White;
				}

				aBuilder.InsertCell();
				aBuilder.Write(property.NetRent);

				if (propCount % 2 == 0)
				{
					firstHalfTable.Rows[RowCount].Cells[7].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
				}
				else
				{
					firstHalfTable.Rows[RowCount].Cells[7].CellFormat.Shading.BackgroundPatternColor = Color.White;
				}

				aBuilder.EndRow();
				propCount++;
			}

			aBuilder.EndTable();
			//UpdateFooter();
			aBuilder.InsertBreak(AWord.BreakType.PageBreak);

			
			font.Color = Color.White;
			AWord.Tables.Table secondHalfTable = aBuilder.StartTable();
			aBuilder.InsertCell();
			aBuilder.Write(" ");
			secondHalfTable.Rows[0].Cells[0].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
			aBuilder.InsertCell();
			aBuilder.Write("OPEX");
			secondHalfTable.Rows[0].Cells[1].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
			aBuilder.InsertCell();
			aBuilder.Write("Taxes");
			secondHalfTable.Rows[0].Cells[2].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
			aBuilder.InsertCell();
			aBuilder.Write("Électricité");
			secondHalfTable.Rows[0].Cells[3].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
			aBuilder.InsertCell();
			aBuilder.Write(" Total des frais additionnels");
			secondHalfTable.Rows[0].Cells[4].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
			aBuilder.InsertCell();
			aBuilder.Write("Loyer Brut");
			secondHalfTable.Rows[0].Cells[5].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
			aBuilder.InsertCell();
			aBuilder.Write("Stationnement");
			secondHalfTable.Rows[0].Cells[6].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
			aBuilder.InsertCell();
			aBuilder.Write("Commentaires");
			secondHalfTable.Rows[0].Cells[7].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;

			aBuilder.EndRow();
			

			propCount = 1;
			foreach (Property property in surveyProperties)
			{
				int RowCount = propCount;
				font.Color = Color.White;
				aBuilder.InsertCell();
				aBuilder.Write(propCount.ToString());
				secondHalfTable.Rows[RowCount].Cells[0].CellFormat.Shading.BackgroundPatternColor = Color.DarkBlue;
				aBuilder.Write(propCount.ToString());

				font.Color = Color.Black;
				aBuilder.InsertCell();
				aBuilder.Write(property.OperationCosts);

				if (propCount % 2 == 0)
				{
					secondHalfTable.Rows[RowCount].Cells[1].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
				}
				else
				{
					secondHalfTable.Rows[RowCount].Cells[1].CellFormat.Shading.BackgroundPatternColor = Color.White;
				}

				aBuilder.InsertCell();
				aBuilder.Write(property.Taxes);

				if (propCount % 2 == 0)
				{
					secondHalfTable.Rows[RowCount].Cells[2].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
				}
				else
				{
					secondHalfTable.Rows[RowCount].Cells[2].CellFormat.Shading.BackgroundPatternColor = Color.White;
				}

				aBuilder.InsertCell();
				aBuilder.Write(property.Energy);

				if (propCount % 2 == 0)
				{
					secondHalfTable.Rows[RowCount].Cells[3].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
				}
				else
				{
					secondHalfTable.Rows[RowCount].Cells[3].CellFormat.Shading.BackgroundPatternColor = Color.White;
				}

				aBuilder.InsertCell();
				aBuilder.Write(property.TotalAdditionalRent);

				if (propCount % 2 == 0)
				{
					secondHalfTable.Rows[RowCount].Cells[4].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
				}
				else
				{
					secondHalfTable.Rows[RowCount].Cells[4].CellFormat.Shading.BackgroundPatternColor = Color.White;
				}

				aBuilder.InsertCell();
				aBuilder.Write(property.GrossRent);

				if (propCount % 2 == 0)
				{
					secondHalfTable.Rows[RowCount].Cells[5].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
				}
				else
				{
					secondHalfTable.Rows[RowCount].Cells[5].CellFormat.Shading.BackgroundPatternColor = Color.White;
				}

				aBuilder.InsertCell();
				aBuilder.Write(property.Parking);

				if (propCount % 2 == 0)
				{
					secondHalfTable.Rows[RowCount].Cells[6].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
				}
				else
				{
					secondHalfTable.Rows[RowCount].Cells[6].CellFormat.Shading.BackgroundPatternColor = Color.White;
				}

				aBuilder.InsertCell();
				aBuilder.Write(property.Comments);

				if (propCount % 2 == 0)
				{
					secondHalfTable.Rows[RowCount].Cells[7].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
				}
				else
				{
					secondHalfTable.Rows[RowCount].Cells[7].CellFormat.Shading.BackgroundPatternColor = Color.White;
				}

				aBuilder.EndRow();
				propCount++;
			}

			aBuilder.EndTable();
			}

			
			
			//UpdateFooter();
			aBuilder.InsertBreak(AWord.BreakType.PageBreak);
		}
		
		
		public void GenerateSiteDetailsTables()
		{

			int propCount = 1;
			
			if (!frenchMode)
			{
				foreach (Property property in surveyProperties)
				{
					AWord.Font font = aBuilder.Font;
					font.Size = 24;
					font.Color = Color.Black;
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.ParagraphFormat.StyleIdentifier = AWord.StyleIdentifier.Heading2;
					aBuilder.Writeln("Site " + propCount.ToString() + "\n" + property.Address);
					aBuilder.ParagraphFormat.StyleIdentifier = AWord.StyleIdentifier.Normal;
					font.Size = 16;


					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					AWord.Tables.Table propertyTable = aBuilder.StartTable();
					aBuilder.CellFormat.TopPadding = 0;
					aBuilder.CellFormat.BottomPadding = 0;
					aBuilder.InsertCell();

					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Administrator/Agency");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.Landlord);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Rentable Area (ft\xB2)");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.RentableArea);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Term");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.Term);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Occupancy");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.Occupancy);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Incentive");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.Incentives);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Net Rental Rate");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.NetRent);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Operating Costs");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.OperationCosts);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Taxes");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.Taxes);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Energy");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.Energy);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Total Additional Rent");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.TotalAdditionalRent);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Font.Bold = true;
					aBuilder.Write("Gross Rent");
					aBuilder.Font.Bold = false;

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.GrossRent);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Parking");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.Parking);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Comments");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.Comments);

					aBuilder.EndRow();
					aBuilder.EndTable();

					string buildingImageFile = surveyGeneralImages.FirstOrDefault(s => s.Contains(property.Address));
					aBuilder.InsertImage(buildingImageFile).Height = 150;
					//UpdateFooter();
					aBuilder.InsertBreak(AWord.BreakType.PageBreak);



					foreach (string file in surveyFloorPlans.FindAll(s => s.Contains(property.Address)))
					{
						aBuilder.InsertImage(file).Height = 300;
					}
					//UpdateFooter();
					aBuilder.InsertBreak(AWord.BreakType.PageBreak);

					foreach (string file in surveyGeneralImages.FindAll(s => s.Contains(property.Address)))
					{
						aBuilder.InsertImage(file).Height = 300;
					}
					propCount++;
					//UpdateFooter();
					aBuilder.InsertBreak(AWord.BreakType.PageBreak);
				}
			}
			else
			{
				foreach (Property property in surveyProperties)
				{
					AWord.Font font = aBuilder.Font;
					font.Size = 24;
					font.Color = Color.Black;
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.ParagraphFormat.StyleIdentifier = AWord.StyleIdentifier.Heading1;
					aBuilder.Writeln("Site " + propCount.ToString() + "\n" + property.Address);
					aBuilder.ParagraphFormat.StyleIdentifier = AWord.StyleIdentifier.Normal;
					font.Size = 16;


					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					AWord.Tables.Table propertyTable = aBuilder.StartTable();
					aBuilder.CellFormat.TopPadding = 0;
					aBuilder.CellFormat.BottomPadding = 0;
					aBuilder.InsertCell();

					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Propriétaire / Administrateur");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.Landlord);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Superficie de l'étage");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.RentableArea);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Terme du bail");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.Term);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Disponibilité");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.Occupancy);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Disponibilité");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.Incentives);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Taux de location net");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.NetRent);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("OPEX");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.OperationCosts);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Taxes");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.Taxes);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Électricité");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.Energy);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Total des frais additionnels");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.TotalAdditionalRent);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Taux de location brut");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.GrossRent);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Stationnement");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.Parking);

					aBuilder.EndRow();

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Left;
					aBuilder.Write("Commentaires");

					aBuilder.InsertCell();
					aBuilder.ParagraphFormat.Alignment = AWord.ParagraphAlignment.Right;
					aBuilder.Write(property.Comments);

					aBuilder.EndRow();
					aBuilder.EndTable();

					string buildingImageFile = surveyGeneralImages.FirstOrDefault(s => s.Contains(property.Address));
					aBuilder.InsertImage(buildingImageFile).Height = 150;
					//UpdateFooter();
					aBuilder.InsertBreak(AWord.BreakType.PageBreak);



					foreach (string file in surveyFloorPlans.FindAll(s => s.Contains(property.Address)))
					{
						aBuilder.InsertImage(file).Height = 300;
					}
					//UpdateFooter();
					aBuilder.InsertBreak(AWord.BreakType.PageBreak);

					foreach (string file in surveyGeneralImages.FindAll(s => s.Contains(property.Address)))
					{
						aBuilder.InsertImage(file).Height = 300;
					}
					propCount++;
					//UpdateFooter();
					aBuilder.InsertBreak(AWord.BreakType.PageBreak);
				}
			}

			
		}
	}
}
