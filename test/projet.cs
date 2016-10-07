/*
 * Crée par SharpDevelop.
 * Utilisateur: ApplicaSuite
 * Date: 16/12/2014
 * Heure: 10:17
 *
 * Pour changer ce modèle utiliser Outils | Options | Codage | Editer les en-têtes standards.
 */

using System;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Data.Common;
using System.Xml;
using System.Text;
using System.Security.Cryptography;


namespace RESAPPLICA
{


	//-----------------------------------------------------------------------------------------------------

	/// <summary>
	/// Gestion de projet commune à tous les modules :
	/// Initialisation projet - Création des tables -
	/// Construction plan maintenance et bilans
	/// </summary>

	public class projet
	{

		private string nomProjet;
		private DataSet dsFichierProjet;
		private DataSet dsProjet;
		private DataSet dsSettings;
		private string nomFichierProjet; // fichier projet
		private string nomCompletFichierDefaultProjet; // fichier defaultProjet.xml
		private string nomFichierBDD; // fichier .mdb
		private string nomFichierDonnees; // fichier .dst
		private string nomFichierSchema; // fichier .rbd
		private string nomFichierXml; // fichier .xml
		private string nomFichierSettingsXml; // fichier .settings.xml
		private string nomCompletFichierDefaultSettingsXml; // fichier defaultSettings.xml

		private string _extension_projet;
		private string _type_projet;
		private string _extension_schema;
		private string _auteur;
		private string _date_creation;
		private string _date_modification;
		private string _indice;

		private string[] tableauParametres;
		private string[] tableauOptions;
		private bool[] tableauProjet;

		private string nomCompletFichierProjet = String.Empty;
		private string nomCompletFichierBDD = String.Empty;
		private string nomCompletFichierDonnees;
		private string nomCompletFichierSchema;
		private string nomCompletFichierXml = String.Empty;
		private string nomCompletFichierSettingsXml = String.Empty;

		private string nomCheminAccesProjet = String.Empty;
		private string nomBaseOperations = String.Empty;
		private string nomCheminBaseOperations = String.Empty;
		private string chaineConnectionProjet = String.Empty;


		public projet(string nom, DataSet ds)
		{
			this.nomProjet = nom;
			this.dsProjet = ds;

			this.nomFichierProjet = this.nomProjet + ".proj.xml";
			this.nomFichierBDD = this.nomProjet + ".mdb";
			this.nomFichierDonnees = this.nomProjet + ".dst";
			this.nomFichierSchema = this.nomProjet + "." + _extension_schema;
			this.nomFichierXml = this.nomProjet + ".xml";
			this.nomFichierSettingsXml = this.nomProjet + ".settings.xml";

			this.nomCheminAccesProjet = Application.StartupPath + "\\PROJETS\\";

			this.nomCompletFichierProjet = nomCheminAccesProjet + nomFichierProjet;
			this.nomCompletFichierBDD = nomCheminAccesProjet + nomFichierBDD;
			this.nomCompletFichierDonnees = nomCheminAccesProjet + nomFichierDonnees;
			this.nomCompletFichierSchema = nomCheminAccesProjet + nomFichierSchema;
			this.nomCompletFichierXml = nomCheminAccesProjet + nomFichierXml;
			this.nomCompletFichierSettingsXml = nomCheminAccesProjet + nomFichierSettingsXml;

			this.nomCompletFichierDefaultSettingsXml = Application.StartupPath + "\\DefaultSettings.xml";
			this.nomCompletFichierDefaultProjet = Application.StartupPath + "\\DefaultProj.xml";

			this.chaineConnectionProjet = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + nomCompletFichierBDD;
			this.nomBaseOperations = "BASE MAINTENANCE_OPERATIONS ET FMD.mdb";
			this.nomCheminBaseOperations = Application.StartupPath + "\\BASE MAINTENANCE_OPERATIONS ET FMD.mdb";

			dsFichierProjet = new DataSet();
			dsProjet = new DataSet();
			dsSettings = new DataSet();

		}

		public string afficheNom{
			get {return nomProjet;}
		}

		public DataSet afficheDsFichierProjet{
			get {return dsFichierProjet;}
		}

		public DataSet afficheDsProjet{
			get {return dsProjet;}
		}

		public DataSet afficheDsSettings{
			get {return dsSettings;}
		}

		public string afficheNomCompletFichierProjet{
			get {return nomCompletFichierProjet;}
		}

		public string afficheNomCompletFichierXml{
			get{return nomCompletFichierXml;}
		}

		public string afficheNomFichierBDD{
			get {return nomFichierBDD;}
		}

		public string afficheNomCompletFichierBDD{
			get {return nomCompletFichierBDD;}
		}

		public string afficheNomFichierXml{
			get {return nomFichierXml;}
		}

		public string afficheNomCompletFichierSettingsXml{
			get{return nomCompletFichierSettingsXml;}
		}

		public string afficheNomBase{
			get {return nomBaseOperations;}
		}

		public string afficheNomCheminBase{
			get {return nomCheminBaseOperations;}
		}

		public string extension_projet{
			get {return _extension_projet;}
			set {this._extension_projet = value;}
		}

		public string extension_schema{
			get {return _extension_schema;}
			set {this._extension_schema = value;}
		}

		public string type_projet{
			get {return _type_projet;}
			set {this._type_projet = value;}
		}

		public string auteur{
			get {return _auteur;}
			set {this._auteur = value;}
		}

		public string date_creation{
			get {return _date_creation;}
			set {this._date_creation = value;}
		}

		public string date_modification{
			get {return _date_modification;}
			set {this._date_modification = value;}
		}

		public string indice{
			get {return _indice;}
			set {this._indice = value;}
		}

		public liste_equipements listeEquipements()
		{
			liste_equipements l1 = new liste_equipements(dsProjet.Tables[0]);
			return l1;
		}

		public liste_plan planMaintenance()
		{
			liste_plan l1 = new liste_plan(dsProjet.Tables[1]);
			return l1;
		}

		public void modifieNom(string _nom)
		{
			this.nomProjet = _nom;
		}

		public void modifieNomFichierProjet(string _nomFichier)
		{
			this.nomFichierProjet = _nomFichier;
		}

		public string afficheParam(int n){
			return tableauParametres[n];
		}

		public string afficheOption(int n){
			return tableauOptions[n];
		}

		public string[] afficheTabParam{
			get {return tableauParametres;}
		}

		public string[] afficheTabOptions{
			get {return tableauOptions;}
		}

		public bool[] afficheTabProj{
			get {return tableauProjet;}
		}

		public bool afficheBool(int n)
		{
			return tableauProjet[n];
		}

		public void modifieBool(int n, bool value)
		{
			this.tableauProjet[n] = value;
		}

		public void modifieParam(int n, string _param)
		{
			this.tableauParametres[n] = _param;
		}

		public void modifieOption(int n, string _option)
		{
			this.tableauOptions[n] = _option;
		}

		public void modifieTabOptions(string[] _tab)
		{
			this.tableauOptions = _tab;
		}

		public void modifieTabParam(string[] _tab)
		{
			this.tableauParametres = _tab;
		}

		public void initialiseTabParam()
		{
			this.tableauParametres = new string[50];
			DataRow dr = dsSettings.Tables[0].Rows[0];
			for(int i = 0; i < tableauParametres.Length; i++){
				this.tableauParametres[i] = dr[i].ToString();
			}
		}

		public void initialiseTabOptions()
		{
			this.tableauOptions = new string[20];
			DataRow dr = dsSettings.Tables[1].Rows[0];
			for(int i = 0; i < tableauOptions.Length; i++){
				this.tableauOptions[i] = dr[i].ToString();
			}
		}

		public void initialiseNomsFichiers(string projectName)
		{
			this.nomCompletFichierProjet = nomCheminAccesProjet + projectName + ".proj.xml";
			this.nomCompletFichierBDD = nomCheminAccesProjet + projectName + ".mdb";
			this.nomCompletFichierDonnees = nomCheminAccesProjet + projectName + ".dst";
			this.nomCompletFichierSchema = nomCheminAccesProjet + projectName + "." + _extension_schema;
			this.nomCompletFichierXml = nomCheminAccesProjet + projectName + ".xml";
			this.nomCompletFichierSettingsXml = nomCheminAccesProjet + projectName + ".settings.xml";
			this.chaineConnectionProjet = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + nomCompletFichierBDD;
		}

		private string[] lireFichierTexte(string chemin, int nbLignes)
		{
			string[] tab = new string[nbLignes];
			StreamReader fichier = File.OpenText(chemin);
			int i = 0;
			while(fichier.Peek() >= 0) {
				string line = fichier.ReadLine();
				string[] valeurs = line.Split(new Char[] {';'});
				tab[i] = valeurs[1];
				i++;
			}
			fichier.Close();
			return tab;
		}

		public void modifieDsProjet(DataSet _ds)
		{
			this.dsProjet.Clear();
			this.dsProjet = _ds;
		}

		public void ajouteDsProjet(DataSet _ds)
		{
			this.dsProjet = _ds;
		}

		public void initialiseTabProjet()
		{
			this.tableauProjet = new bool[15];
			if(dsProjet.Tables.Contains("Table Liste EQUIPEMENTS") == true){tableauProjet[8] = true;}else{tableauProjet[8] = false;}
			if(dsProjet.Tables.Contains("Table Frequences") == true){tableauProjet[5] = true;}else{tableauProjet[5] = false;}
			if(dsProjet.Tables.Contains("Table Plan de maintenance [par REPERE_EQUIPEMENT]") == true){tableauProjet[9] = true;}else{tableauProjet[9] = false;}
			if(dsProjet.Tables.Contains("Table Charge maintenance preventive") == true){tableauProjet[10] = true;}else{tableauProjet[10] = false;}
			if(dsProjet.Tables.Contains("Table Charge maintenance preventive (consolidée)") == true){tableauProjet[11] = true;}else{tableauProjet[11] = false;}
			if(dsProjet.Tables.Contains("Table Charge maintenance corrective") == true){tableauProjet[12] = true;}else{tableauProjet[12] = false;}
			if(dsProjet.Tables.Contains("Table Liste CABLES") == true){tableauProjet[13] = true;}else{tableauProjet[13] = false;}
			tableauProjet[0] = false;
			tableauProjet[1] = false;
			tableauProjet[2] = false;
			tableauProjet[3] = false;
			tableauProjet[4] = false;
			tableauProjet[6] = false;
			tableauProjet[7] = false;
			tableauProjet[14] = false;
		}

		// définition de tabProj (boolean)

		// tabProj[0]	existence fichier BDD .mdb
		// tabProj[1]	existence fichier .dst
		// tabProj[2]	existence fichier .rbd
		// tabProj[3]	existence fichier .xml
		// tabProj[4]
		// tabProj[5]	existence Table Frequences dans dsProjet
		// tabProj[6]	projet modifié
		// tabProj[7]	projet enregistré
		// tabProj[8]	existence Table Liste EQUIPEMENTS dans dsProjet
		// tabProj[9]	existence Table Plan de maintenance [par REPERE_EQUIPEMENT] dans dsProjet
		// tabProj[10]	existence Table Charge maintenance preventive dans dsProjet
		// tabProj[11]	existence Table Charge maintenance preventive (consolidée) dans dsProjet
		// tabProj[12]	existence Table Charge maintenance corrective dans dsProjet
		// tabProj[13]	existence Table Liste CABLES
		// tabProj[14]

		private void initialiserDsProjetXml()
		{
			dsProjet.ReadXml(nomCompletFichierXml);
		}

		private void initialiserDsFichierProjet()
		{
			if(!File.Exists(nomCompletFichierProjet)){
				dsFichierProjet.ReadXml(nomCompletFichierDefaultProjet);
			}else{
				dsFichierProjet.ReadXml(nomCompletFichierProjet);
			}
		}

		private void initialiserDsSettingsXml()
		{
			if(!File.Exists(nomCompletFichierSettingsXml)){
				dsSettings.ReadXml(nomCompletFichierDefaultSettingsXml);
			}else{
				dsSettings.ReadXml(nomCompletFichierSettingsXml);
			}
		}

		public void chargerDefaultprojet()
		{
			string nomFichierDefaultProj = Application.StartupPath + "\\Defaultproj.xml";
			dsFichierProjet.ReadXml(nomFichierDefaultProj);
		}

		public void chargerDefaultsettings()
		{
			string nomFichierDefaultSettings = Application.StartupPath + "\\DefaultSettings.xml";
			dsSettings.ReadXml(nomFichierDefaultSettings);
		}

		public string chaineParametres()
		{
			string chaine = String.Empty;
			for(int i = 0; i < this.afficheTabParam.Length; i++){
				chaine = chaine + this.dsSettings.Tables[0].Columns[i] + "\t\t" + this.afficheTabParam[i] + "\n";
			}
			return chaine;
		}

		public string chaineOptions()
		{
			string chaine = String.Empty;
			for(int i = 0; i < this.afficheTabOptions.Length; i++){
				chaine = chaine + this.dsSettings.Tables[1].Columns[i] + "\t\t" + this.afficheTabOptions[i] + "\n";
			}
			return chaine;
		}

		public void initialiser()
		{
			initialiserDsFichierProjet();
			initialiserDsSettingsXml();
			initialiserDsProjetXml();
			initialiseTabParam();
			initialiseTabOptions();
			initialiseTabProjet();
		}

		public string option_15()
		{
			string selection = "CreationBD_projet";
			return afficheDsSettings.Tables["OPTIONS"].Rows[0][selection].ToString();
		}

		public string option_16()
		{
			string selection = "EnregistrerChargeMP";
			return afficheDsSettings.Tables["OPTIONS"].Rows[0][selection].ToString();
		}

		public string option_17()
		{
			string selection = "EnregistrerChargeMC";
			return afficheDsSettings.Tables["OPTIONS"].Rows[0][selection].ToString();
		}

		public void mettreAJourDsFichierProjet()
		{
			afficheDsFichierProjet.Tables[0].Rows[0]["NOM_PROJET"] = nomProjet;
			afficheDsFichierProjet.Tables[0].Rows[0]["TYPE_PROJET"] = _type_projet;
			afficheDsFichierProjet.Tables[0].Rows[0]["AUTEUR"] = _auteur;
			afficheDsFichierProjet.Tables[0].Rows[0]["DATE_CREATION"] = _date_creation;
			afficheDsFichierProjet.Tables[0].Rows[0]["DATE_MODIFICATION"] = _date_modification;
			afficheDsFichierProjet.Tables[0].Rows[0]["INDICE"] = _indice;
		}

		public void mettreAJourDsSettings()
		{
			// mise à jour Paramètres
			for(int i = 0; i < tableauParametres.Length; i++){
				if(tableauParametres[i] != null){
					afficheDsSettings.Tables["PARAMETRES"].Rows[0][i] = tableauParametres[i];
				}else{
					afficheDsSettings.Tables["PARAMETRES"].Rows[0][i] = DBNull.Value;
				}
			}
			// mise à jour Options
			for(int i = 0; i < tableauOptions.Length; i++){
				if(tableauOptions[i] != null){
					afficheDsSettings.Tables["OPTIONS"].Rows[0][i] = tableauOptions[i];
				}else{
					afficheDsSettings.Tables["OPTIONS"].Rows[0][i] = DBNull.Value;;
				}
			}
			// mise à jour Taux horaires


		}

		public void enregistrer()
		{
//			enregistrerParametres();
			enregistrerFichierProjet();
			enregistrerDsSettingsXml();
			enregistrerDsProjetXml();
		}

		public string[] nomsParametres()
		{
			int dimension = dsSettings.Tables["PARAMETRES"].Columns.Count;
			string[] tab = new string[dimension];
			int i = 0;
			foreach(DataColumn c in dsSettings.Tables["PARAMETRES"].Columns){
				tab[i] = c.ColumnName.ToString();
				i++;
			}
			return tab;
		}

		private void enregistrerParametres()
		{
			if (File.Exists(nomCompletFichierProjet)){File.Delete(nomCompletFichierProjet);}
			StreamWriter fichier = new StreamWriter(nomCompletFichierProjet);
			string[] tab = nomsParametres();
			int _tailleParametres_1 = 23;
			for(int i = 0; i < _tailleParametres_1; i++){
				string line = tab[i] + ";" + afficheParam(i);
				fichier.WriteLine(line);
			}
			fichier.Close();
		}

		private void enregistrerFichierProjet()
		{
			dsFichierProjet.WriteXml(@nomCompletFichierProjet, XmlWriteMode.WriteSchema);
		}

		public void enregistrerDsProjetXml()
		{
			dsProjet.WriteXml(@nomCompletFichierXml, XmlWriteMode.WriteSchema);
		}

		public void enregistrerDsSettingsXml()
		{
			dsSettings.WriteXml(@nomCompletFichierSettingsXml, XmlWriteMode.WriteSchema);
		}

		public DataTable listeEquipementsPerimetre(string selection, string ordreTri)
		{
			DataTable dt = creerTableListeEquipements();
			DataTable dtEquipements = afficheDsProjet.Tables["Table Liste EQUIPEMENTS"];
			DataRow[] dr = dtEquipements.Select(selection, ordreTri);

			for(int i = 0; i < dr.Length; i++){
				DataRow r = dt.NewRow();
				r["NUM"] = dr[i]["NUM"];
				r["QUANTITE"] = dr[i]["QUANTITE"];
				r["TYPE_EQUIPEMENT"] = dr[i]["TYPE_EQUIPEMENT"];
				r["REPERE_EQUIPEMENT"] = dr[i]["REPERE_EQUIPEMENT"];
				r["NOM_SYSTEME"] = dr[i]["NOM_SYSTEME"];
				r["DESIGNATION"] = dr[i]["DESIGNATION"];
				r["REPERE_SECTEUR"] = dr[i]["REPERE_SECTEUR"];
				r["REPERE_BATIMENT"] = dr[i]["REPERE_BATIMENT"];
				r["REPERE_EMPLACEMENT"] = dr[i]["REPERE_EMPLACEMENT"];
				r["LOCALISATION"] = dr[i]["LOCALISATION"];
				r["MARQUE_TYPE"] = dr[i]["MARQUE_TYPE"];
				r["ANNEE_MISE_EN_SERVICE"] = dr[i]["ANNEE_MISE_EN_SERVICE"];
				r["PARAMETRE_A"] = dr[i]["PARAMETRE_A"];
				r["PARAMETRE_B"] = dr[i]["PARAMETRE_B"];
				r["PARAMETRE_C"] = dr[i]["PARAMETRE_C"];
				r["PARAMETRE_D"] = dr[i]["PARAMETRE_D"];
				r["CLASSE_G"] = dr[i]["CLASSE_G"];
				r["APPARTENANCE_PERIMETRE"] = dr[i]["APPARTENANCE_PERIMETRE"];
				r["CARACTERE"] = dr[i]["CARACTERE"];
				r["CHAMP_PARAM_1"] = dr[i]["CHAMP_PARAM_1"];
				r["CHAMP_PARAM_2"] = dr[i]["CHAMP_PARAM_2"];
				r["CHAMP_PARAM_3"] = dr[i]["CHAMP_PARAM_3"];
				r["CHAMP_PARAM_4"] = dr[i]["CHAMP_PARAM_4"];
				r["CHAMP_PARAM_5"] = dr[i]["CHAMP_PARAM_5"];
				dt.Rows.Add(r);
			}
			return dt;
		}

		public double numeroEquipementSuivant()
		{
			double numero = 1;
			DataTable dt = afficheDsProjet.Tables["Table Liste EQUIPEMENTS"];
			double[] tab = creerListeIndex(dt);
			if(tab.Length != 0){
				numero = tab[tab.Length -1] + 1;
			}
			return numero;
		}

		public string repereEquipementSuivant()
		{
			string result = String.Empty;
			DataTable dt = afficheDsProjet.Tables["Table Liste EQUIPEMENTS"];
			double rang = numeroEquipementSuivant();
			if(rang  < 10){
				result = "REP_000" + rang.ToString();
			}else{
				if(rang  < 100){
					result = "REP_00" + rang.ToString();
				}else{
					if(rang  < 1000){
						result = "REP_0" + rang.ToString();
					}else{
						if(rang  < 10000){
							result = "REP_" + rang.ToString();
						}
					}
				}
			}
			bool exist = verifierExistenceRepereEquipement(result);
			switch (exist){
				case true :
					result = result + "A";
				break;
				case false :
				break;
			}

			return result;
		}

		private double[] creerListeIndex(DataTable dt)
		{
			int lignes = dt.Rows.Count;
			double[] tab = new double[lignes];
			if(lignes != 0){
				for(int i = 0; i < lignes; i++){
					if(dt.Rows[i].RowState != DataRowState.Deleted){
						if(dt.Rows[i]["NUM"].ToString() != ""){
							tab[i] = double.Parse(dt.Rows[i]["NUM"].ToString());
						}
					}
				}
				Array.Sort(tab);
			}
			return tab;
		}

		private string[] creerListeReperes()
		{
			DataTable dt = afficheDsProjet.Tables["Table Liste EQUIPEMENTS"];
			int lignes = dt.Rows.Count;
			string[] tab = new string[lignes];

			for(int i = 0; i < lignes; i++){
				if(dt.Rows[i].RowState != DataRowState.Deleted){
					tab[i] = dt.Rows[i]["REPERE_EQUIPEMENT"].ToString();
				// Array.Sort(tab);
				}
			}
			return tab;
		}

		private bool verifierExistenceRepereEquipement(string rep)
		{
			bool result = true; // true si le repere equipement existe
			string[] tab = creerListeReperes();
			int ind = Array.IndexOf(tab, rep);
			if(ind == -1){result = false;} else{result = true;}
			return result;
		}

		public void ajouterEquipement(equipement eqpt)
		{
			DataTable dt = afficheDsProjet.Tables["Table Liste EQUIPEMENTS"];
			DataRow dr = dt.NewRow();
			dr["NUM"] = double.Parse(eqpt.afficheCar(29));
			if(eqpt.afficheCar(11) == ""){dr["QUANTITE"] = 1;}else{dr["QUANTITE"] = double.Parse(eqpt.afficheCar(11));}
			dr["TYPE_EQUIPEMENT"] = eqpt.afficheCar(0);
			dr["REPERE_EQUIPEMENT"] = eqpt.afficheRepere;
			if(eqpt.afficheCar(2) == ""){dr["NOM_SYSTEME"] = "(non affecté)";}else{dr["NOM_SYSTEME"] = eqpt.afficheCar(2);}
			dr["DESIGNATION"] = eqpt.afficheCar(1);
			if(eqpt.afficheCar(4) == ""){dr["REPERE_SECTEUR"] = "(non défini)";}else{dr["REPERE_SECTEUR"] = eqpt.afficheCar(4);}
			dr["REPERE_BATIMENT"] = eqpt.afficheCar(5);
			dr["REPERE_EMPLACEMENT"] = eqpt.afficheCar(6);
			dr["LOCALISATION"] = eqpt.afficheCar(7);
			dr["MARQUE_TYPE"] = eqpt.afficheCar(3);
			dr["ANNEE_MISE_EN_SERVICE"] = eqpt.afficheCar(8);
			dr["PARAMETRE_A"] = eqpt.afficheParamEqpt(0);
			dr["PARAMETRE_B"] = eqpt.afficheParamEqpt(1);
			dr["PARAMETRE_C"] = eqpt.afficheParamEqpt(2);
			dr["PARAMETRE_D"] = eqpt.afficheParamEqpt(3);
			dr["CLASSE_G"] = int.Parse(eqpt.afficheCar(9));
			if(eqpt.afficheCar(17) == ""){dr["APPARTENANCE_PERIMETRE"] = "OUI";}else{dr["APPARTENANCE_PERIMETRE"] = eqpt.afficheCar(17);}
			if(eqpt.afficheCar(11) == "1"){dr["CARACTERE"] = "IND";}else{dr["CARACTERE"] = "REG";}
			// dr["CHAMP_PARAM_1"] = eqpt.afficheCar(12);
			// dr["CHAMP_PARAM_2"] = eqpt.afficheCar(13);
			// dr["CHAMP_PARAM_3"] = eqpt.afficheCar(14);
			// dr["CHAMP_PARAM_4"] = eqpt.afficheCar(15);
			// dr["CHAMP_PARAM_5"] = eqpt.afficheCar(16);
			dt.Rows.Add(dr);
		}

		public DataTable creerTableListeEquipements()
		{
			DataTable dt = new DataTable("Table Liste EQUIPEMENTS");
			dt.Columns.Add("NUM", System.Type.GetType("System.Double"));
			dt.Columns.Add("REPERE_EQUIPEMENT", System.Type.GetType("System.String"));
			dt.Columns.Add("NOM_SYSTEME", System.Type.GetType("System.String"));
			dt.Columns.Add("DESIGNATION", System.Type.GetType("System.String"));
			dt.Columns.Add("TYPE_EQUIPEMENT", System.Type.GetType("System.String"));
			dt.Columns.Add("MARQUE_TYPE", System.Type.GetType("System.String"));
			dt.Columns.Add("REPERE_SECTEUR", System.Type.GetType("System.String"));
			dt.Columns.Add("REPERE_BATIMENT", System.Type.GetType("System.String"));
			dt.Columns.Add("REPERE_EMPLACEMENT", System.Type.GetType("System.String"));
			dt.Columns.Add("LOCALISATION", System.Type.GetType("System.String"));
			dt.Columns.Add("ANNEE_MISE_EN_SERVICE", System.Type.GetType("System.String"));
			dt.Columns.Add("PARAMETRE_A", System.Type.GetType("System.String"));
			dt.Columns.Add("PARAMETRE_B", System.Type.GetType("System.String"));
			dt.Columns.Add("PARAMETRE_C", System.Type.GetType("System.String"));
			dt.Columns.Add("PARAMETRE_D", System.Type.GetType("System.String"));
			dt.Columns.Add("CLASSE_G", System.Type.GetType("System.Double"));
			dt.Columns.Add("CARACTERE", System.Type.GetType("System.String"));
			dt.Columns.Add("QUANTITE", System.Type.GetType("System.Double"));
			dt.Columns.Add("CHAMP_PARAM_1", System.Type.GetType("System.String"));
			dt.Columns.Add("CHAMP_PARAM_2", System.Type.GetType("System.String"));
			dt.Columns.Add("CHAMP_PARAM_3", System.Type.GetType("System.String"));
			dt.Columns.Add("CHAMP_PARAM_4", System.Type.GetType("System.String"));
			dt.Columns.Add("CHAMP_PARAM_5", System.Type.GetType("System.String"));
			dt.Columns.Add("APPARTENANCE_PERIMETRE", System.Type.GetType("System.String"));
			return dt;
		}

		public DataTable extractionListeEquipements(string selection, string ordreTri){

			DataTable dt = creerTableListeEquipements();
			DataTable dtEquipements = afficheDsProjet.Tables["Table Liste EQUIPEMENTS"];
			DataRow[] dr = dtEquipements.Select(selection, ordreTri);

			for(int i = 0; i < dr.Length; i++){
				DataRow r = dt.NewRow();
				r["NUM"] = dr[i]["NUM"];
				r["QUANTITE"] = dr[i]["QUANTITE"];
				r["TYPE_EQUIPEMENT"] = dr[i]["TYPE_EQUIPEMENT"];
				r["REPERE_EQUIPEMENT"] = dr[i]["REPERE_EQUIPEMENT"];
				r["NOM_SYSTEME"] = dr[i]["NOM_SYSTEME"];
				r["DESIGNATION"] = dr[i]["DESIGNATION"];
				r["REPERE_SECTEUR"] = dr[i]["REPERE_SECTEUR"];
				r["REPERE_BATIMENT"] = dr[i]["REPERE_BATIMENT"];
				r["REPERE_EMPLACEMENT"] = dr[i]["REPERE_EMPLACEMENT"];
				r["LOCALISATION"] = dr[i]["LOCALISATION"];
				r["MARQUE_TYPE"] = dr[i]["MARQUE_TYPE"];
				r["ANNEE_MISE_EN_SERVICE"] = dr[i]["ANNEE_MISE_EN_SERVICE"];
				r["PARAMETRE_A"] = dr[i]["PARAMETRE_A"];
				r["PARAMETRE_B"] = dr[i]["PARAMETRE_B"];
				r["PARAMETRE_C"] = dr[i]["PARAMETRE_C"];
				r["PARAMETRE_D"] = dr[i]["PARAMETRE_D"];
				r["CLASSE_G"] = dr[i]["CLASSE_G"];
				r["APPARTENANCE_PERIMETRE"] = dr[i]["APPARTENANCE_PERIMETRE"];
				r["CARACTERE"] = dr[i]["CARACTERE"];
				r["CHAMP_PARAM_1"] = dr[i]["CHAMP_PARAM_1"];
				r["CHAMP_PARAM_2"] = dr[i]["CHAMP_PARAM_2"];
				r["CHAMP_PARAM_3"] = dr[i]["CHAMP_PARAM_3"];
				r["CHAMP_PARAM_4"] = dr[i]["CHAMP_PARAM_4"];
				r["CHAMP_PARAM_5"] = dr[i]["CHAMP_PARAM_5"];
				dt.Rows.Add(r);
			}
			return dt;
		}

		public DataTable creerTableListeEquipements(string nomMapping)
		{
			DataTable dt = new DataTable(nomMapping);
			dt.Columns.Add("NUM", System.Type.GetType("System.Double"));
			dt.Columns.Add("REPERE_EQUIPEMENT", System.Type.GetType("System.String"));
			dt.Columns.Add("NOM_SYSTEME", System.Type.GetType("System.String"));
			dt.Columns.Add("DESIGNATION", System.Type.GetType("System.String"));
			dt.Columns.Add("TYPE_EQUIPEMENT", System.Type.GetType("System.String"));
			dt.Columns.Add("MARQUE_TYPE", System.Type.GetType("System.String"));
			dt.Columns.Add("REPERE_SECTEUR", System.Type.GetType("System.String"));
			dt.Columns.Add("REPERE_BATIMENT", System.Type.GetType("System.String"));
			dt.Columns.Add("REPERE_EMPLACEMENT", System.Type.GetType("System.String"));
			dt.Columns.Add("LOCALISATION", System.Type.GetType("System.String"));
			dt.Columns.Add("ANNEE_MISE_EN_SERVICE", System.Type.GetType("System.String"));
			dt.Columns.Add("PARAMETRE_A", System.Type.GetType("System.String"));
			dt.Columns.Add("PARAMETRE_B", System.Type.GetType("System.String"));
			dt.Columns.Add("PARAMETRE_C", System.Type.GetType("System.String"));
			dt.Columns.Add("PARAMETRE_D", System.Type.GetType("System.String"));
			dt.Columns.Add("CLASSE_G", System.Type.GetType("System.Double"));
			dt.Columns.Add("CARACTERE", System.Type.GetType("System.String"));
			dt.Columns.Add("QUANTITE", System.Type.GetType("System.Double"));
			dt.Columns.Add("CHAMP_PARAM_1", System.Type.GetType("System.String"));
			dt.Columns.Add("CHAMP_PARAM_2", System.Type.GetType("System.String"));
			dt.Columns.Add("CHAMP_PARAM_3", System.Type.GetType("System.String"));
			dt.Columns.Add("CHAMP_PARAM_4", System.Type.GetType("System.String"));
			dt.Columns.Add("CHAMP_PARAM_5", System.Type.GetType("System.String"));
			dt.Columns.Add("APPARTENANCE_PERIMETRE", System.Type.GetType("System.String"));
			return dt;
		}

		public DataTable extractionPlanMaintenance(string selection, string ordreTri){

			DataTable dt = creerTablePlanMaintenance();
			DataTable dtPlan = afficheDsProjet.Tables["Table Plan de maintenance [par REPERE_EQUIPEMENT]"];
			DataRow[] dr = dtPlan.Select(selection, ordreTri);

			for(int i = 0; i < dr.Length; i++){
				DataRow r = dt.NewRow();
				r["CODE_OPERATION"] = dr[i]["CODE_OPERATION"];
				r["NOM_OPERATION_MAINTENANCE"] = dr[i]["NOM_OPERATION_MAINTENANCE"];
				r["TYPE_VISITE"] = dr[i]["TYPE_VISITE"];
				r["PERIODICITE"] = dr[i]["PERIODICITE"];
				r["UNITE_PERIODICITE"] = dr[i]["UNITE_PERIODICITE"];
				r["REPERE_EQUIPEMENT"] = dr[i]["REPERE_EQUIPEMENT"];
				r["NOM_SYSTEME"] = dr[i]["NOM_SYSTEME"];
				r["DESIGNATION"] = dr[i]["DESIGNATION"];
				r["TYPE_EQUIPEMENT"] = dr[i]["TYPE_EQUIPEMENT"];
				r["REPERE_SECTEUR"] = dr[i]["REPERE_SECTEUR"];
				r["REPERE_BATIMENT"] = dr[i]["REPERE_BATIMENT"];
				r["REPERE_EMPLACEMENT"] = dr[i]["REPERE_EMPLACEMENT"];
				r["LOCALISATION"] = dr[i]["LOCALISATION"];
				r["QUANTITE"] = dr[i]["QUANTITE"];
				r["CLASSE_G"] = dr[i]["CLASSE_G"];
				r["CHAMP_PARAM_1"] = dr[i]["CHAMP_PARAM_1"];
				r["CHAMP_PARAM_2"] = dr[i]["CHAMP_PARAM_2"];
				r["CHAMP_PARAM_3"] = dr[i]["CHAMP_PARAM_3"];
				r["CHAMP_PARAM_4"] = dr[i]["CHAMP_PARAM_4"];
				r["CHAMP_PARAM_5"] = dr[i]["CHAMP_PARAM_5"];
				dt.Rows.Add(r);
			}
			return dt;
		}

		public DataTable extractionPlanMaintenance2(string selection, string ordreTri){

			DataTable dt = creerTablePlanMaintenance2();
			DataTable dtPlan = afficheDsProjet.Tables["Table Plan de maintenance [par REPERE_EQUIPEMENT]"];
			DataRow[] dr = dtPlan.Select(selection, ordreTri);

			for(int i = 0; i < dr.Length; i++){
				DataRow r = dt.NewRow();
				r["CODE_OPERATION"] = dr[i]["CODE_OPERATION"];
				r["PERIODICITE"] = dr[i]["PERIODICITE"];
				r["UNITE_PERIODICITE"] = dr[i]["UNITE_PERIODICITE"];
				r["REPERE_EQUIPEMENT"] = dr[i]["REPERE_EQUIPEMENT"];
				r["NOM_SYSTEME"] = dr[i]["NOM_SYSTEME"];
				r["DESIGNATION"] = dr[i]["DESIGNATION"];
				r["TYPE_EQUIPEMENT"] = dr[i]["TYPE_EQUIPEMENT"];
				r["REPERE_SECTEUR"] = dr[i]["REPERE_SECTEUR"];
				r["REPERE_BATIMENT"] = dr[i]["REPERE_BATIMENT"];
				r["REPERE_EMPLACEMENT"] = dr[i]["REPERE_EMPLACEMENT"];
				r["LOCALISATION"] = dr[i]["LOCALISATION"];
				r["QUANTITE"] = dr[i]["QUANTITE"];
				r["CLASSE_G"] = dr[i]["CLASSE_G"];
				r["CHAMP_PARAM_1"] = dr[i]["CHAMP_PARAM_1"];
				r["CHAMP_PARAM_2"] = dr[i]["CHAMP_PARAM_2"];
				r["CHAMP_PARAM_3"] = dr[i]["CHAMP_PARAM_3"];
				r["CHAMP_PARAM_4"] = dr[i]["CHAMP_PARAM_4"];
				r["CHAMP_PARAM_5"] = dr[i]["CHAMP_PARAM_5"];
				dt.Rows.Add(r);
			}
			return dt;
		}

		public DataTable creerTablePlanMaintenance()
		{
			DataTable dt = new DataTable("Table Plan de maintenance [par REPERE_EQUIPEMENT]");
			dt.Columns.Add("CODE_OPERATION", System.Type.GetType("System.String"));
			dt.Columns.Add("NOM_OPERATION_MAINTENANCE", System.Type.GetType("System.String"));
			dt.Columns.Add("TYPE_VISITE", System.Type.GetType("System.String"));
			dt.Columns.Add("PERIODICITE", System.Type.GetType("System.String"));
			dt.Columns.Add("UNITE_PERIODICITE", System.Type.GetType("System.String"));
			dt.Columns.Add("REPERE_EQUIPEMENT", System.Type.GetType("System.String"));
			dt.Columns.Add("NOM_SYSTEME", System.Type.GetType("System.String"));
			dt.Columns.Add("DESIGNATION", System.Type.GetType("System.String"));
			dt.Columns.Add("TYPE_EQUIPEMENT", System.Type.GetType("System.String"));
			dt.Columns.Add("REPERE_SECTEUR", System.Type.GetType("System.String"));
			dt.Columns.Add("REPERE_BATIMENT", System.Type.GetType("System.String"));
			dt.Columns.Add("REPERE_EMPLACEMENT", System.Type.GetType("System.String"));
			dt.Columns.Add("LOCALISATION", System.Type.GetType("System.String"));
			dt.Columns.Add("QUANTITE", System.Type.GetType("System.Double"));
			dt.Columns.Add("CLASSE_G", System.Type.GetType("System.Double"));
			dt.Columns.Add("CHAMP_PARAM_1", System.Type.GetType("System.String"));
			dt.Columns.Add("CHAMP_PARAM_2", System.Type.GetType("System.String"));
			dt.Columns.Add("CHAMP_PARAM_3", System.Type.GetType("System.String"));
			dt.Columns.Add("CHAMP_PARAM_4", System.Type.GetType("System.String"));
			dt.Columns.Add("CHAMP_PARAM_5", System.Type.GetType("System.String"));
			return dt;
		}

		public DataTable creerTablePlanMaintenance2()
		{
			DataTable dt = new DataTable("Table Plan de maintenance [par REPERE_EQUIPEMENT]");
			dt.Columns.Add("CODE_OPERATION", System.Type.GetType("System.String"));

			dt.Columns.Add("PERIODICITE", System.Type.GetType("System.String"));
			dt.Columns.Add("UNITE_PERIODICITE", System.Type.GetType("System.String"));
			dt.Columns.Add("REPERE_EQUIPEMENT", System.Type.GetType("System.String"));
			dt.Columns.Add("NOM_SYSTEME", System.Type.GetType("System.String"));
			dt.Columns.Add("DESIGNATION", System.Type.GetType("System.String"));
			dt.Columns.Add("TYPE_EQUIPEMENT", System.Type.GetType("System.String"));
			dt.Columns.Add("REPERE_SECTEUR", System.Type.GetType("System.String"));
			dt.Columns.Add("REPERE_BATIMENT", System.Type.GetType("System.String"));
			dt.Columns.Add("REPERE_EMPLACEMENT", System.Type.GetType("System.String"));
			dt.Columns.Add("LOCALISATION", System.Type.GetType("System.String"));
			dt.Columns.Add("QUANTITE", System.Type.GetType("System.Double"));
			dt.Columns.Add("CLASSE_G", System.Type.GetType("System.Double"));
			dt.Columns.Add("CHAMP_PARAM_1", System.Type.GetType("System.String"));
			dt.Columns.Add("CHAMP_PARAM_2", System.Type.GetType("System.String"));
			dt.Columns.Add("CHAMP_PARAM_3", System.Type.GetType("System.String"));
			dt.Columns.Add("CHAMP_PARAM_4", System.Type.GetType("System.String"));
			dt.Columns.Add("CHAMP_PARAM_5", System.Type.GetType("System.String"));
			return dt;
		}

		public DataTable creerTablePlanMaintenance(string nomMapping)
		{
			DataTable dt = new DataTable(nomMapping);
			dt.Columns.Add("CODE_OPERATION", System.Type.GetType("System.String"));
			dt.Columns.Add("NOM_OPERATION_MAINTENANCE", System.Type.GetType("System.String"));
			dt.Columns.Add("TYPE_VISITE", System.Type.GetType("System.String"));
			dt.Columns.Add("PERIODICITE", System.Type.GetType("System.String"));
			dt.Columns.Add("UNITE_PERIODICITE", System.Type.GetType("System.String"));
			dt.Columns.Add("REPERE_EQUIPEMENT", System.Type.GetType("System.String"));
			dt.Columns.Add("NOM_SYSTEME", System.Type.GetType("System.String"));
			dt.Columns.Add("DESIGNATION", System.Type.GetType("System.String"));
			dt.Columns.Add("TYPE_EQUIPEMENT", System.Type.GetType("System.String"));
			dt.Columns.Add("REPERE_SECTEUR", System.Type.GetType("System.String"));
			dt.Columns.Add("REPERE_BATIMENT", System.Type.GetType("System.String"));
			dt.Columns.Add("REPERE_EMPLACEMENT", System.Type.GetType("System.String"));
			dt.Columns.Add("LOCALISATION", System.Type.GetType("System.String"));
			dt.Columns.Add("QUANTITE", System.Type.GetType("System.Double"));
			dt.Columns.Add("CLASSE_G", System.Type.GetType("System.Double"));
			dt.Columns.Add("CHAMP_PARAM_1", System.Type.GetType("System.String"));
			dt.Columns.Add("CHAMP_PARAM_2", System.Type.GetType("System.String"));
			dt.Columns.Add("CHAMP_PARAM_3", System.Type.GetType("System.String"));
			dt.Columns.Add("CHAMP_PARAM_4", System.Type.GetType("System.String"));
			dt.Columns.Add("CHAMP_PARAM_5", System.Type.GetType("System.String"));
			// dt.Columns.Add("APPARTENANCE_PERIMETRE", System.Type.GetType("System.String"));
			return dt;
		}

		public DataTable creerTableListeOperations()
		{
			DataTable dt = new DataTable("Table Liste Operations");
			dt.Columns.Add("CODE_OPERATION", System.Type.GetType("System.String"));
			dt.Columns.Add("NOM_OPERATION_MAINTENANCE", System.Type.GetType("System.String"));
			dt.Columns.Add("TYPE_VISITE", System.Type.GetType("System.String"));
			dt.Columns.Add("PERIODICITE", System.Type.GetType("System.String"));
			dt.Columns.Add("UNITE_PERIODICITE", System.Type.GetType("System.String"));
			return dt;
		}

		public int nombreEnregistrementsListeTrouves(visu vis)
		{
			string chaineSelectionListe =  vis.chaineSelection();
			string chaineTri =  vis.chaineTri;
			DataTable dataTable1 = extractionListeEquipements(chaineSelectionListe, chaineTri);
			return dataTable1.Rows.Count;
		}

		public int nombreEnregistrementsPlanTrouves(visu vis)
		{
			string chaineSelectionListe =  vis.chaineSelection();
			string chaineTri =  vis.chaineTri;
			DataTable dataTable1 = extractionPlanMaintenance(chaineSelectionListe, chaineTri);
			return dataTable1.Rows.Count;
		}

		public DataTable listeOperations(string repere, string ordreTri)
		{
			DataTable dt = creerTableListeOperations();
			string selection = "(REPERE_EQUIPEMENT = '" + repere + "')";

			DataTable dtPlanMaintenance = afficheDsProjet.Tables["Table Plan de maintenance [par REPERE_EQUIPEMENT]"];
			DataRow[] dr = dtPlanMaintenance.Select(selection, ordreTri);

			for(int i = 0; i < dr.Length; i++){
				DataRow r = dt.NewRow();
				r["CODE_OPERATION"] = dr[i]["CODE_OPERATION"];
				r["NOM_OPERATION_MAINTENANCE"] = dr[i]["NOM_OPERATION_MAINTENANCE"];
				r["TYPE_VISITE"] = dr[i]["TYPE_VISITE"];
				r["PERIODICITE"] = dr[i]["PERIODICITE"];
				r["UNITE_PERIODICITE"] = dr[i]["UNITE_PERIODICITE"];
				dt.Rows.Add(r);
			}
			return dt;
		}

		public DataTable creerTableFrequences()
		{
			DataTable dt = new DataTable("Table Frequences");
			dt.Columns.Add("REPERE_EQUIPEMENT", System.Type.GetType("System.String"));
			dt.Columns.Add("CODE_OPERATION", System.Type.GetType("System.String"));
			dt.Columns.Add("PERIODICITE", System.Type.GetType("System.String"));
			dt.Columns.Add("UNITE_PERIODICITE", System.Type.GetType("System.String"));
			return dt;
		}

		public void initialiserTableTauxHoraires()
		{
			DataTable dt= afficheDsSettings.Tables["TAUX HORAIRES"].Copy();
			dt.TableName = "Table TAUX_HORAIRES";
			afficheDsProjet.Tables.Add(dt);
		}

		public DataTable creerTableBilan()
		{
			DataTable dt = new DataTable("Table Bilan");
			dt.Columns.Add("REPERE", System.Type.GetType("System.String"));
			dt.Columns.Add("DESIGNATION", System.Type.GetType("System.String"));
			dt.Columns.Add("VALEUR_0", System.Type.GetType("System.Double")); // classe_G 0
			dt.Columns.Add("VALEUR_1", System.Type.GetType("System.Double")); // classe_G 1
			dt.Columns.Add("VALEUR_2", System.Type.GetType("System.Double")); // classe_G 2
			return dt;
		}

		public DataTable listeEquipementsPlan(DataTable dt0)
		{
			string nomTable = "Liste EQUIPEMENTS du PLAN";
			DataTable dt = selectDistinct(nomTable, dt0, "REPERE_EQUIPEMENT");
			return dt;
		}

		private DataTable selectDistinct(string nomTable, DataTable tableSource, string champ)
		{
		   DataTable dt = new DataTable(nomTable);
		   dt.Columns.Add(champ, tableSource.Columns[champ].DataType);
		   object LastValue = null;
		   foreach (DataRow dr in tableSource.Select("", champ)){
			   if(LastValue == null || !(DataColumn.Equals(LastValue, dr[champ]))){
				  LastValue = dr[champ];
				  dt.Rows.Add(new object[]{LastValue});
			   }
		   }
		   return dt;
		}

		public int rechercherIndexTable(string repere)
		{
			int index = 0;
			DataTable dt = dsProjet.Tables["Table Liste EQUIPEMENTS"];
			int j = 0;
			foreach(DataRow dr in dt.Rows){
				if((string)dr["REPERE_EQUIPEMENT"] == repere){
					index = j;
				}
				j++;
			}
			return index;
		}

		public string[] listeSystemes()
		{
			DataTable dt = selectDistinct("", afficheDsProjet.Tables["Table Liste EQUIPEMENTS"], "NOM_SYSTEME");
			string[] tab = new string[dt.Rows.Count];
			int i = 0;
			foreach(DataRow r in dt.Rows){
				tab[i] = r["NOM_SYSTEME"].ToString();
				i++;
			}
			return tab;
		}

		public string[] listeBatiments()
		{
			DataTable dt = selectDistinct("", afficheDsProjet.Tables["Table Liste EQUIPEMENTS"], "REPERE_BATIMENT");
			string[] tab = new string[dt.Rows.Count];
			int i = 0;
			foreach(DataRow r in dt.Rows){
				tab[i] = r["REPERE_BATIMENT"].ToString();
				i++;
			}
			return tab;
		}

		public string[] listeSecteurs() // sans les champs vides
		{
			DataTable dt = selectDistinct_blank("", afficheDsProjet.Tables["Table Liste EQUIPEMENTS"], "REPERE_SECTEUR");
			string[] tab = new string[dt.Rows.Count];
			int i = 0;
			foreach(DataRow r in dt.Rows){
				tab[i] = r["REPERE_SECTEUR"].ToString();
				i++;
			}
			return tab;
		}

		public string[] listeSecteurs_complet() // avec les champs vides
		{
			DataTable dt = selectDistinct("", afficheDsProjet.Tables["Table Liste EQUIPEMENTS"], "REPERE_SECTEUR");
			string[] tab = new string[dt.Rows.Count];
			int i = 0;
			foreach(DataRow r in dt.Rows){
				tab[i] = r["REPERE_SECTEUR"].ToString();
				if(tab[i] == String.Empty){tab[i] = "(non défini)";}
				i++;
			}
			return tab;
		}

		private DataTable selectDistinct_blank(string nomTable, DataTable tableSource, string champ)
		{
		   	DataTable dt = new DataTable(nomTable);
			dt.Columns.Add(champ, tableSource.Columns[champ].DataType);
			string valeur = String.Empty;
			foreach (DataRow dr in tableSource.Select("", champ)){
				if(!(DataColumn.Equals(valeur, dr[champ].ToString()))){
					valeur = dr[champ].ToString();
					DataRow r = dt.NewRow();
					r[champ] = valeur;
					dt.Rows.Add(r);
				}
			}
		   return dt;
		}

		private string[] champsLocalisation(string loc)
		{
			string[] res = new string[4];
			string[] valeurs = loc.Split(new Char[] {'/'});
			if(valeurs.Length > 2){
					res[0] = valeurs[0];
					res[1] = valeurs[1];
					res[2] = valeurs[2];
			}
			else{
				if(valeurs.Length > 1){
					res[0] = valeurs[0];
					res[1] = valeurs[1];
				}
				else{
					res[0] = valeurs[0];
				}
			}
			res[3] = loc;
			return res;
		}

		public string[,] quantitesParSecteur()
		{
			string[] secteurs = listeSecteurs();
			string[,] tab = new string[secteurs.Length,2];
			for(int i = 0; i < secteurs.Length; i++){
				string selection = "REPERE_SECTEUR ='" + secteurs[i] + "'";
				DataRow[] dr = afficheDsProjet.Tables["Table Liste EQUIPEMENTS"].Select(selection);
				tab[i,0] = secteurs[i];
				double q = 0;
				for(int j = 0; j < dr.Length; j++){
					q = q + double.Parse(dr[j]["QUANTITE"].ToString());
				}
				tab[i,1] = q.ToString();
			}
			return tab;
		}

		public double quantitesParTypeParSecteur(string type, string secteur)
		{
			string selection = "TYPE_EQUIPEMENT='" + type + "' AND REPERE_SECTEUR='" + secteur + "'";
			DataRow[] dr = afficheDsProjet.Tables["Table Liste EQUIPEMENTS"].Select(selection);
			double quantite = 0;
			for(int j = 0; j < dr.Length; j++){
				double q = double.Parse(dr[j]["QUANTITE"].ToString());
				quantite = quantite + q;
			}
			return quantite;
		}

		public string[,] quantitesParSysteme()
		{
			string[] systemes = listeSystemes();
			string[,] tab = new string[systemes.Length,2];
			for(int i = 0; i < systemes.Length; i++){
				string strSQL = "NOM_SYSTEME ='" + systemes[i] + "'";
				DataRow[] dr = afficheDsProjet.Tables["Table Liste EQUIPEMENTS"].Select(strSQL);
				tab[i,0] = systemes[i];
				double q = 0;
				for(int j = 0; j < dr.Length; j++){
					q = q + double.Parse(dr[j]["QUANTITE"].ToString());
				}
				tab[i,1] = q.ToString();
			}
			return tab;
		}

		public double quantitesParTypeParSysteme(string type, string systeme)
		{
			string selection = "TYPE_EQUIPEMENT='" + type + "' AND NOM_SYSTEME='" + systeme + "'";
			DataRow[] dr = afficheDsProjet.Tables["Table Liste EQUIPEMENTS"].Select(selection);
			double quantite = 0;
			for(int j = 0; j < dr.Length; j++){
				double q = double.Parse(dr[j]["QUANTITE"].ToString());
				quantite = quantite + q;
			}
			return quantite;
		}

		public void modifierEquipement(int index, equipement eqpt)
		{
			DataTable dt = afficheDsProjet.Tables["Table Liste EQUIPEMENTS"];
			dt.Rows[index]["TYPE_EQUIPEMENT"] = eqpt.afficheCar(0);
			dt.Rows[index]["DESIGNATION"] = eqpt.afficheCar(1);
			dt.Rows[index]["NOM_SYSTEME"] = eqpt.afficheCar(2);
			dt.Rows[index]["MARQUE_TYPE"] = eqpt.afficheCar(3);
			dt.Rows[index]["REPERE_SECTEUR"] = eqpt.afficheCar(4);
			dt.Rows[index]["REPERE_BATIMENT"] = eqpt.afficheCar(5);
			dt.Rows[index]["REPERE_EMPLACEMENT"] = eqpt.afficheCar(6);
			dt.Rows[index]["LOCALISATION"] = eqpt.afficheCar(7);
			dt.Rows[index]["ANNEE_MISE_EN_SERVICE"] = eqpt.afficheCar(8);
			dt.Rows[index]["CLASSE_G"] = eqpt.afficheCar(9);
			dt.Rows[index]["CARACTERE"] = eqpt.afficheCar(10);
			dt.Rows[index]["QUANTITE"] = eqpt.afficheCar(11);
			dt.Rows[index]["APPARTENANCE_PERIMETRE"]= eqpt.afficheCar(17);
			dt.Rows[index]["PARAMETRE_A"] = eqpt.afficheParamEqpt(0);
			dt.Rows[index]["PARAMETRE_B"] = eqpt.afficheParamEqpt(1);
			dt.Rows[index]["PARAMETRE_C"] = eqpt.afficheParamEqpt(2);
			dt.Rows[index]["PARAMETRE_D"] = eqpt.afficheParamEqpt(3);
		}

		public double charge_MPREV(string systeme, double classe, string champ)
		{
			double res = 0;
			string[] listeReperes = listeEquipementsParSysteme(systeme, classe);
			switch(champ){
			case "MO" :
				for(int i = 0; i < listeReperes.Length; i++){
					res = res + totalTempsMPREVparEquipement(listeReperes[i]);
				}
			break;
			case "CONS" :
				for(int i = 0; i < listeReperes.Length; i++){
					res = res + totalConsommablesMPREVparEquipement(listeReperes[i]);
				}
			break;
			case "PR" :
				for(int i = 0; i < listeReperes.Length; i++){
					res = res + totalPiecesMPREVparEquipement(listeReperes[i]);
				}
			break;
			case "ST" :
				for(int i = 0; i < listeReperes.Length; i++){
					res = res + totalSTMPREVparEquipement(listeReperes[i]);
				}
			break;
			default :
				res = 0;
			break;
			}
			return res;
		}

		public string[] listeEquipementsParSysteme(string syst)
		{
			string selection = "NOM_SYSTEME ='" + syst + "'";
			DataRow[] dr = afficheDsProjet.Tables["Table Liste EQUIPEMENTS"].Select(selection);
			string[] tab = new string[dr.Length];
			for(int i = 0; i < dr.Length; i++){
				tab[i] = dr[i]["REPERE_EQUIPEMENT"].ToString();
			}
			return tab;
		}

		public string[] listeEquipementsParSysteme(string syst, double classe)
		{
			string selection = "NOM_SYSTEME ='" + syst + "' AND CLASSE_G ='" + classe.ToString() + "'";
			DataRow[] dr = afficheDsProjet.Tables["Table Liste EQUIPEMENTS"].Select(selection);
			string[] tab = new string[dr.Length];
			for(int i = 0; i < dr.Length; i++){
				tab[i] = dr[i]["REPERE_EQUIPEMENT"].ToString();
			}
			return tab;
		}

		public double totalTempsMPREVparEquipement(string repere)
		{
			double totalTempsMPREV = 0;
			DataTable dtChargePREVCons = afficheDsProjet.Tables["Table Charge maintenance preventive (consolidée)"];
			string selection = "REPERE_EQUIPEMENT ='" + repere + "'";
			DataRow[] dr1 = dtChargePREVCons.Select(selection);

			for(int i = 0; i < dr1.Length; i++){
				totalTempsMPREV = totalTempsMPREV + double.Parse(dr1[i]["TEMPS_TOTAL"].ToString());
			}
			return totalTempsMPREV;
		}

		public double totalConsommablesMPREVparEquipement(string repere)
		{
			double totalConsoMPREV = 0;
			DataTable dtChargePREVCons = afficheDsProjet.Tables["Table Charge maintenance preventive (consolidée)"];
			string selection = "REPERE_EQUIPEMENT ='" + repere + "'";
			DataRow[] dr1 = dtChargePREVCons.Select(selection);

			for(int i = 0; i < dr1.Length; i++){
				totalConsoMPREV = totalConsoMPREV + double.Parse(dr1[i]["COUT_CONS_TOTAL"].ToString());
			}
			return totalConsoMPREV;
		}

		private double totalConsommablesMCORparEquipement(string repere)
		{
			double totalConsoMCOR = 0;
			DataTable dtChargeCOR = afficheDsProjet.Tables["Table Charge maintenance corrective"];
			string selection = "REPERE_EQUIPEMENT ='" + repere + "'";
			DataRow[] dr1 = dtChargeCOR.Select(selection);

			for(int i = 0; i < dr1.Length; i++){
				totalConsoMCOR = totalConsoMCOR + double.Parse(dr1[i]["COUT_CONS_TOTAL"].ToString());
			}
			return totalConsoMCOR;
		}

		private double totalPiecesMPREVparEquipement(string repere)
		{
			double totalPiecesMPREV = 0;
			DataTable dtChargePREVCons = afficheDsProjet.Tables["Table Charge maintenance preventive (consolidée)"];
			string selection = "REPERE_EQUIPEMENT ='" + repere + "'";
			DataRow[] dr1 = dtChargePREVCons.Select(selection);

			for(int i = 0; i < dr1.Length; i++){
				totalPiecesMPREV = totalPiecesMPREV + double.Parse(dr1[i]["COUT_PIECES_TOTAL"].ToString());
			}
			return totalPiecesMPREV;
		}

		private double totalPiecesMCORparEquipement(string repere)
		{
			double totalPiecesMCOR = 0;
			DataTable dtChargeCOR = afficheDsProjet.Tables["Table Charge maintenance corrective"];
			string selection = "REPERE_EQUIPEMENT ='" + repere + "'";
			DataRow[] dr1 = dtChargeCOR.Select(selection);

			for(int i = 0; i < dr1.Length; i++){
				totalPiecesMCOR = totalPiecesMCOR + double.Parse(dr1[i]["COUT_MAT_TOTAL"].ToString());
			}
			return totalPiecesMCOR;
		}

		private double totalSTMPREVparEquipement(string repere)
		{
			double totalSTMPREV = 0;
			DataTable dtChargePREVCons = afficheDsProjet.Tables["Table Charge maintenance preventive (consolidée)"];
			string selection = "REPERE_EQUIPEMENT ='" + repere + "'";
			DataRow[] dr1 = dtChargePREVCons.Select(selection);

			for(int i = 0; i < dr1.Length; i++){
				totalSTMPREV = totalSTMPREV + double.Parse(dr1[i]["COUT_ST_TOTAL"].ToString());
			}
			return totalSTMPREV;
		}

		private double totalSTMCORparEquipement(string repere)
		{
			double totalSTMCOR = 0;
			DataTable dtChargeCOR = afficheDsProjet.Tables["Table Charge maintenance corrective"];
			string selection = "REPERE_EQUIPEMENT ='" + repere + "'";
			DataRow[] dr1 = dtChargeCOR.Select(selection);

			for(int i = 0; i < dr1.Length; i++){
				totalSTMCOR = totalSTMCOR + double.Parse(dr1[i]["COUT_ST_TOTAL"].ToString());
			}
			return totalSTMCOR;
		}

		public double totalTempsMPREV()
		{
			double totalTempsMPREV = 0;
			DataTable dtChargePREVCons = afficheDsProjet.Tables["Table Charge maintenance preventive (consolidée)"];

			for(int i = 0; i < dtChargePREVCons.Rows.Count - 2; i++){
				totalTempsMPREV = totalTempsMPREV + double.Parse(dtChargePREVCons.Rows[i]["TEMPS_TOTAL"].ToString());
			}
			return totalTempsMPREV;
		}

		public double totalTempsMCOR()
		{
			double totalTempsMCOR = 0;
			DataTable dtChargeCORCons = afficheDsProjet.Tables["Table Charge maintenance corrective"];

			for(int i = 0; i < dtChargeCORCons.Rows.Count - 2; i++){
				totalTempsMCOR = totalTempsMCOR + double.Parse(dtChargeCORCons.Rows[i]["TEMPS_TOTAL"].ToString());
			}
			return totalTempsMCOR;
		}

		public double totalOperationsMPREV()
		{
			double totalOperationsMPREV = 0;
			DataTable dtChargePREVCons = afficheDsProjet.Tables["Table Charge maintenance preventive (consolidée)"];

			for(int i = 0; i < dtChargePREVCons.Rows.Count - 2; i++){
				totalOperationsMPREV = totalOperationsMPREV + double.Parse(dtChargePREVCons.Rows[i]["N_OPERATIONS"].ToString());
			}
			return totalOperationsMPREV;
		}

		public double[] totalOperations(string systeme){

			double[] tab = new double[4];
			string[] listeReperes = listeEquipementsParSysteme(systeme);

			for(int i = 0; i < listeReperes.Length; i++){
				tab[0] = tab[0] + nOperationsMPREVparEquipement(listeReperes[i]);
				tab[1] = tab[1] + totalTempsMPREVparEquipement(listeReperes[i]);
				tab[2] = tab[2] + nIntMCORparEquipement(listeReperes[i]);
				tab[3] = tab[3] + totalTempsMCORparEquipement(listeReperes[i]);
			}
			return tab;
		}

		public double totalVisitesMPREV()
		{
			double totalVisitesMPREV = 0;
			DataTable dtChargePREVCons = afficheDsProjet.Tables["Table Charge maintenance preventive (consolidée)"];

			for(int i = 0; i < dtChargePREVCons.Rows.Count - 2; i++){
				totalVisitesMPREV = totalVisitesMPREV + double.Parse(dtChargePREVCons.Rows[i]["N_VISITES"].ToString());
			}
			return totalVisitesMPREV;
		}

		public double totalIntMCOR(){

			double totalIntMCOR = 0;
			DataTable dtChargeCOR = afficheDsProjet.Tables["Table Charge maintenance corrective"];

			for(int i = 0; i < dtChargeCOR.Rows.Count - 2; i++){
				totalIntMCOR = totalIntMCOR + double.Parse(dtChargeCOR.Rows[i]["N_INT"].ToString());
			}
			return totalIntMCOR;
		}

		public double tempsMoyenPREV_parOperation()
		{
			double tempsMoyenPREV = totalTempsMPREV() / totalOperationsMPREV();
			return tempsMoyenPREV;
		}

		public double tempsMoyenPREV_parVisite()
		{
			double tempsMoyenPREV = totalTempsMPREV() / totalVisitesMPREV();
			return tempsMoyenPREV;
		}

		public double nombreCodesOperationsPlan()
		{
			double nCodesOperations = 0;
			DataTable dt = afficheDsProjet.Tables["Table Plan de maintenance [par REPERE_EQUIPEMENT]"];
			nCodesOperations = dt.Rows.Count;
			return nCodesOperations;
		}

		private double nOperationsMPREVparEquipement(string repere)
		{
			double  nOperationsMPREV = 0;
			DataTable dtChargePREVCons = afficheDsProjet.Tables["Table Charge maintenance preventive (consolidée)"];
			string selection = "REPERE_EQUIPEMENT ='" + repere + "'";
			DataRow[] dr1 = dtChargePREVCons.Select(selection);

			for(int i = 0; i < dr1.Length; i++){
				nOperationsMPREV = nOperationsMPREV + double.Parse(dr1[i]["N_OPERATIONS"].ToString());
			}
			return nOperationsMPREV;
		}

		public double nVisitesMPREVparEquipement(string repere)
		{
			double  nVisitesMPREV = 0;
			DataTable dtChargePREVCons = afficheDsProjet.Tables["Table Charge maintenance preventive (consolidée)"];
			string selection = "REPERE_EQUIPEMENT ='" + repere + "'";
			DataRow[] dr1 = dtChargePREVCons.Select(selection);

			for(int i = 0; i < dr1.Length; i++){
				nVisitesMPREV = nVisitesMPREV + double.Parse(dr1[i]["N_VISITES"].ToString());
			}
			return nVisitesMPREV;
		}

		public double totalTempsMCORparEquipement(string repere)
		{
			double totalTempsMCOR = 0;
			DataTable dtChargeCOR = afficheDsProjet.Tables["Table Charge maintenance corrective"];
			string selection = "REPERE_EQUIPEMENT ='" + repere + "'";
			DataRow[] dr1 = dtChargeCOR.Select(selection);

			for(int i = 0; i < dr1.Length; i++){
				totalTempsMCOR = totalTempsMCOR + double.Parse(dr1[i]["TEMPS_TOTAL"].ToString());
			}
			return totalTempsMCOR;
		}

		private double nIntMCORparEquipement(string repere)
		{
			double   nIntMCOR = 0;
			DataTable dtChargeCOR = afficheDsProjet.Tables["Table Charge maintenance corrective"];
			string selection = "REPERE_EQUIPEMENT ='" + repere + "'";
			DataRow[] dr1 = dtChargeCOR.Select(selection);

			for(int i = 0; i < dr1.Length; i++){
				 nIntMCOR =  nIntMCOR + double.Parse(dr1[i]["N_INT"].ToString());
			}
			return nIntMCOR;
		}

		public double nombreTotalEquipements()
		{
			double res = 0;
			for(int i = 0; i < afficheDsProjet.Tables["Table Liste EQUIPEMENTS"].Rows.Count; i++){
				res = res + double.Parse(afficheDsProjet.Tables["Table Liste EQUIPEMENTS"].Rows[i]["QUANTITE"].ToString());
			}
			return res;
		}

		public double nombreTotalEquipements(int classe)
		{
			double res = 0;
			string selection = "CLASSE_G ='" + classe.ToString() + "'";
			DataRow[] dr1 = afficheDsProjet.Tables["Table Liste EQUIPEMENTS"].Select(selection);
			for(int i = 0; i < dr1.Length; i++){
				 res =  res + double.Parse(dr1[i]["QUANTITE"].ToString());
			}
			return res;
		}

		public double charge_MCOR(string systeme, double classe, string champ)
		{
			double res = 0;
			string[] listeReperes = listeEquipementsParSysteme(systeme, classe);
			switch(champ){
				case "MO" :
					for(int i = 0; i < listeReperes.Length; i++){
						res = res + totalTempsMCORparEquipement(listeReperes[i]);
					}
				break;
				case "CONS" :
					for(int i = 0; i < listeReperes.Length; i++){
						res = res + totalConsommablesMCORparEquipement(listeReperes[i]);
					}
				break;
				case "PR" :
					for(int i = 0; i < listeReperes.Length; i++){
						res = res + totalPiecesMCORparEquipement(listeReperes[i]);
					}
				break;
				case "ST" :
					for(int i = 0; i < listeReperes.Length; i++){
						res = res + totalSTMCORparEquipement(listeReperes[i]);
					}
				break;
				default :
					res = 0;
				break;
			}
			return res;
		}

		public void enregistrerTableListeEquipementsBDD()
		{
			DataTable dt = dsProjet.Tables["Table Liste EQUIPEMENTS"];
			string nomTable = "TABLE_LISTE_EQUIPEMENTS";
			string repere, nomSysteme, designation, type, annee, marque, secteur, batiment, emplacement, localisation, caractere;
			string champ1, champ2, champ3, champ4, champ5, appartenance;
			double num, parametreA, parametreB, parametreC, parametreD, classeG, quantite;

			// définition des paramètres
			OleDbParameter paramNum = new OleDbParameter("@NUM", OleDbType.Double);
			OleDbParameter paramRepere = new OleDbParameter("@REPERE_EQUIPEMENT", OleDbType.VarChar, 50);
			OleDbParameter paramNomSysteme = new OleDbParameter("@NOM_SYSTEME", OleDbType.VarChar, 50);
			OleDbParameter paramDesignation = new OleDbParameter("@DESIGNATION", OleDbType.VarChar, 50);
			OleDbParameter paramType = new OleDbParameter("@TYPE_EQUIPEMENT", OleDbType.VarChar, 255);
			OleDbParameter paramMarque = new OleDbParameter("@MARQUE_TYPE", OleDbType.VarChar, 50);
			OleDbParameter paramSecteur = new OleDbParameter("@REPERE_SECTEUR", OleDbType.VarChar, 50);
			OleDbParameter paramBatiment = new OleDbParameter("@REPERE_BATIMENT", OleDbType.VarChar, 50);
			OleDbParameter paramEmplacement = new OleDbParameter("@REPERE_EMPLACEMENT", OleDbType.VarChar, 50);
			OleDbParameter paramLocalisation = new OleDbParameter("@xxx", OleDbType.VarChar, 255);
			OleDbParameter paramAnnee = new OleDbParameter("@ANNEE_MISE_EN_SERVICE", OleDbType.VarChar, 50);
			OleDbParameter paramParA = new OleDbParameter("@PARAMETRE_A", OleDbType.Double);
			OleDbParameter paramParB = new OleDbParameter("@PARAMETRE_B", OleDbType.Double);
			OleDbParameter paramParC = new OleDbParameter("@PARAMETRE_C", OleDbType.Double);
			OleDbParameter paramParD = new OleDbParameter("@PARAMETRE_D", OleDbType.Double);
			OleDbParameter paramClasseG = new OleDbParameter("@CLASSE_G", OleDbType.Double);
			OleDbParameter paramCaractere = new OleDbParameter("@CARACTERE", OleDbType.VarChar, 3);
			OleDbParameter paramQuantite = new OleDbParameter("@QUANTITE", OleDbType.Double);
			OleDbParameter paramChamp1 = new OleDbParameter("@CHAMP_PARAM_1", OleDbType.VarChar, 50);
			OleDbParameter paramChamp2 = new OleDbParameter("@CHAMP_PARAM_2", OleDbType.VarChar, 50);
			OleDbParameter paramChamp3 = new OleDbParameter("@CHAMP_PARAM_3", OleDbType.VarChar, 50);
			OleDbParameter paramChamp4 = new OleDbParameter("@CHAMP_PARAM_4", OleDbType.VarChar, 50);
			OleDbParameter paramChamp5 = new OleDbParameter("@CHAMP_PARAM_5", OleDbType.VarChar, 50);
			OleDbParameter paramAppartenance = new OleDbParameter("@APPARTENANCE_PERIMETRE", OleDbType.VarChar, 3);

			OleDbConnection connection1 = new OleDbConnection(chaineConnectionProjet);
			connection1.Open();

			String strEnreg = string.Format("INSERT INTO {0}(NUM, REPERE_EQUIPEMENT, NOM_SYSTEME, DESIGNATION, TYPE_EQUIPEMENT, MARQUE_TYPE, REPERE_SECTEUR, REPERE_BATIMENT, REPERE_EMPLACEMENT, LOCALISATION, ANNEE_MISE_EN_SERVICE, PARAMETRE_A, PARAMETRE_B, PARAMETRE_C, PARAMETRE_D, CLASSE_G, CARACTERE, QUANTITE, CHAMP_PARAM_1, CHAMP_PARAM_2, CHAMP_PARAM_3, CHAMP_PARAM_4, CHAMP_PARAM_5, APPARTENANCE_PERIMETRE) VALUES({1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24})",
			nomTable, paramNum.ParameterName, paramRepere.ParameterName, paramNomSysteme.ParameterName, paramDesignation.ParameterName, paramType.ParameterName, paramMarque.ParameterName, paramSecteur.ParameterName, paramBatiment.ParameterName, paramEmplacement.ParameterName, paramLocalisation.ParameterName, paramAnnee.ParameterName, paramParA.ParameterName, paramParB.ParameterName, paramParC.ParameterName, paramParD.ParameterName, paramClasseG.ParameterName, paramCaractere.ParameterName, paramQuantite.ParameterName, paramChamp1.ParameterName, paramChamp2.ParameterName, paramChamp3.ParameterName, paramChamp4.ParameterName, paramChamp5.ParameterName, paramAppartenance.ParameterName);

			OleDbCommand command1 = new OleDbCommand(strEnreg.ToString(), connection1);

			command1.Parameters.Add(paramNum);
			command1.Parameters.Add(paramRepere);
			command1.Parameters.Add(paramNomSysteme);
			command1.Parameters.Add(paramDesignation);
			command1.Parameters.Add(paramType);
			command1.Parameters.Add(paramMarque);
			command1.Parameters.Add(paramSecteur);
			command1.Parameters.Add(paramBatiment);
			command1.Parameters.Add(paramEmplacement);
			command1.Parameters.Add(paramLocalisation);
			command1.Parameters.Add(paramAnnee);
			command1.Parameters.Add(paramParA);
			command1.Parameters.Add(paramParB);
			command1.Parameters.Add(paramParC);
			command1.Parameters.Add(paramParD);
			command1.Parameters.Add(paramClasseG);
			command1.Parameters.Add(paramCaractere);
			command1.Parameters.Add(paramQuantite);
			command1.Parameters.Add(paramChamp1);
			command1.Parameters.Add(paramChamp2);
			command1.Parameters.Add(paramChamp3);
			command1.Parameters.Add(paramChamp4);
			command1.Parameters.Add(paramChamp5);
			command1.Parameters.Add(paramAppartenance);

			for(int i = 0; i < dt.Rows.Count; i++){

				DataRow dr = dt.Rows[i];
				num = (double)(i + 1);
				repere = dr["REPERE_EQUIPEMENT"].ToString();
				nomSysteme = dr["NOM_SYSTEME"].ToString();
				if(nomSysteme == ""){nomSysteme = "(non affecté)";}
				designation = dr["DESIGNATION"].ToString();
				type = (string) dr["TYPE_EQUIPEMENT"];
				annee = dr["ANNEE_MISE_EN_SERVICE"].ToString();
				marque = dr["MARQUE_TYPE"].ToString();
				secteur = dr["REPERE_SECTEUR"].ToString();
				batiment = dr["REPERE_BATIMENT"].ToString();
				emplacement = dr["REPERE_EMPLACEMENT"].ToString();

				localisation = chaineLocalisation(secteur, batiment, emplacement);

				if(dr["PARAMETRE_A"].ToString() != ""){parametreA = double.Parse(dr["PARAMETRE_A"].ToString());}else{parametreA = 2;}
				if(dr["PARAMETRE_B"].ToString() != ""){parametreB = double.Parse(dr["PARAMETRE_B"].ToString());}else{parametreB = 2;}
				if(dr["PARAMETRE_C"].ToString() != ""){parametreC = double.Parse(dr["PARAMETRE_C"].ToString());}else{parametreC = 2;}
				if(dr["PARAMETRE_D"].ToString() != ""){parametreD = double.Parse(dr["PARAMETRE_D"].ToString());}else{parametreD = 2;}

				if(dr["CLASSE_G"].ToString() != ""){classeG = double.Parse(dr["CLASSE_G"].ToString());}else{classeG = 0;}

				caractere = dr["CARACTERE"].ToString();
				quantite = double.Parse(dr["QUANTITE"].ToString());
				champ1 = dr["CHAMP_PARAM_1"].ToString();
				champ2 = dr["CHAMP_PARAM_2"].ToString();
				champ3 = dr["CHAMP_PARAM_3"].ToString();
				champ4 = dr["CHAMP_PARAM_4"].ToString();
				champ5 = dr["CHAMP_PARAM_5"].ToString();
				appartenance = dr["APPARTENANCE_PERIMETRE"].ToString();
				if(appartenance == ""){appartenance = "OUI";}

				// affectation des paramètres
				paramNum.Value = num;
				paramRepere.Value = repere;
				paramNomSysteme.Value = nomSysteme;
				paramDesignation.Value = designation;
				paramType.Value = type;
				paramMarque.Value = marque;
				paramSecteur.Value = secteur;
				paramBatiment.Value = batiment;
				paramEmplacement.Value = emplacement;
				paramLocalisation.Value = localisation;
				paramAnnee.Value = annee;
				paramParA.Value = parametreA;
				paramParB.Value = parametreB;
				paramParC.Value = parametreC;
				paramParD.Value = parametreD;
				paramClasseG.Value = classeG;
				paramCaractere.Value = caractere;
				paramQuantite.Value = quantite;
				paramChamp1.Value = champ1;
				paramChamp2.Value = champ2;
				paramChamp3.Value = champ3;
				paramChamp4.Value = champ4;
				paramChamp5.Value = champ5;
				paramAppartenance.Value = appartenance;

				command1.ExecuteNonQuery();
			}
			connection1.Close();
		}

		public void enregistrerTableListeSystemesBDD()
		{
			string nomSysteme = "(non affecté)";
			string codeSysteme = String.Empty;
			string[] listeSys = listeSystemes();

			// définition des paramètres
			OleDbParameter paramNomSysteme = new OleDbParameter("@NOM_SYSTEME", OleDbType.VarChar, 50);
			OleDbParameter paramCodeSysteme = new OleDbParameter("@CODE_SYSTEME", OleDbType.VarChar, 50);
			OleDbConnection connection1 = new OleDbConnection(chaineConnectionProjet);
			connection1.Open();
			string nomTable = "TABLE_LISTE_SYSTEMES";
			string strEnreg = string.Format("INSERT INTO {0}(NOM_SYSTEME, CODE_SYSTEME) VALUES({1},{2})", nomTable, paramNomSysteme.ParameterName, paramCodeSysteme.ParameterName);

			OleDbCommand command1 = new OleDbCommand(strEnreg.ToString(), connection1);
			command1.Parameters.Add(paramNomSysteme);
			command1.Parameters.Add(paramCodeSysteme);

			for(int i = 0; i < listeSys.Length; i++){
				nomSysteme = listeSys[i];
				codeSysteme = (i + 1).ToString();
				paramNomSysteme.Value = nomSysteme;
				paramCodeSysteme.Value = codeSysteme;
				command1.ExecuteNonQuery();
			}
			connection1.Close();
		}

		public void enregistrerTableFrequencesBDD()
		{
			DataTable dt = afficheDsProjet.Tables["Table Frequences"];
			double num = 0;
			string repere, code, periodicite, unitePeriodicite;
			string nomTable = "TABLE_FREQUENCES";

			OleDbParameter paramNum = new OleDbParameter("@NUM2", OleDbType.Double);
			OleDbParameter paramRepere = new OleDbParameter("@REPERE_EQUIPEMENT", OleDbType.VarChar, 50);
			OleDbParameter paramCode = new OleDbParameter("@CODE_OPERATION", OleDbType.VarChar, 50);
			OleDbParameter paramPeriodicite = new OleDbParameter("@PERIODICITE", OleDbType.Double);
			OleDbParameter paramUnitePeriodicite = new OleDbParameter("@UNITE_PERIODICITE", OleDbType.VarChar, 50);

			OleDbConnection connection1 = new OleDbConnection(chaineConnectionProjet);
			connection1.Open();
			String strEnreg = string.Format("INSERT INTO {0}(NUM2, REPERE_EQUIPEMENT, CODE_OPERATION, PERIODICITE, UNITE_PERIODICITE) VALUES({1},{2},{3},{4},{5})", nomTable, paramNum.ParameterName, paramRepere.ParameterName, paramCode.ParameterName, paramPeriodicite.ParameterName, paramUnitePeriodicite.ParameterName);
			OleDbCommand command1 = new OleDbCommand(strEnreg.ToString(), connection1);

			command1.Parameters.Add(paramNum);
			command1.Parameters.Add(paramRepere);
			command1.Parameters.Add(paramCode);
			command1.Parameters.Add(paramPeriodicite);
			command1.Parameters.Add(paramUnitePeriodicite);

			for(int i = 0; i < dt.Rows.Count; i++){
				DataRow dr = dt.Rows[i];
				num = (double)(i + 1);
				repere = dr["REPERE_EQUIPEMENT"].ToString();
				code = dr["CODE_OPERATION"].ToString();
				periodicite = dr["PERIODICITE"].ToString();
				unitePeriodicite = dr["UNITE_PERIODICITE"].ToString();

				paramNum.Value = num;
				paramRepere.Value = repere;
				paramCode.Value = code;
				paramPeriodicite.Value = periodicite;
				paramUnitePeriodicite.Value = unitePeriodicite;

				command1.ExecuteNonQuery();
			}
			connection1.Close();
		}

		private void creerTableCHARGE_MP_BDD(string nomTable)
		{
			string strEnreg = "CREATE TABLE " + nomTable + "(NUM DOUBLE PRECISION NOT NULL PRIMARY KEY, REPERE_EQUIPEMENT VARCHAR(255), TYPE_EQUIPEMENT VARCHAR(255), QUANTITE DOUBLE PRECISION, TEMPS_NORM DOUBLE PRECISION,"
				+ " COUT_NORM_CONS DOUBLE PRECISION, COUT_NORM_PIECES DOUBLE PRECISION, COUT_NORM_ST DOUBLE PRECISION,"
					+ " N_OPERATIONS DOUBLE PRECISION, N_VISITES DOUBLE PRECISION, TEMPS_TOTAL DOUBLE PRECISION,"
						+ " COUT_CONS_TOTAL DOUBLE PRECISION, COUT_PIECES_TOTAL DOUBLE PRECISION, COUT_ST_TOTAL DOUBLE PRECISION)";
			OleDbConnection connection1 = new OleDbConnection(chaineConnectionProjet);
			OleDbCommand command1 = new OleDbCommand(strEnreg);
			command1.Connection = connection1;
			connection1.Open();
			command1.ExecuteNonQuery();
			connection1.Close();
		}

		public void enregistrerTableChargeMP_BDD()
		{
			DataTable dt = dsProjet.Tables["Table Charge maintenance preventive (consolidée)"];
			double num = 0;
			string repere;
			string typeEquipement;
			double quantite = 0;
			double tempsNorm = 0;
			double coutNormCons = 0;
			double coutNormPieces = 0;
			double coutNormST = 0;
			double Noperations = 0;
			double Nvisites = 0;
			double tempsTotal = 0;
			double coutConsTotal = 0;
			double coutPiecesTotal = 0;
			double coutSTTotal = 0;
			string nomTable = "TABLE_CHARGE_MPREV";

			creerTableCHARGE_MP_BDD(nomTable);

			OleDbParameter paramNum = new OleDbParameter("@NUM", OleDbType.Double);
			OleDbParameter paramRepere = new OleDbParameter("@REPERE_EQUIPEMENT", OleDbType.VarChar, 50);
			OleDbParameter paramTypeEquipement = new OleDbParameter("@TYPE_EQUIPEMENT", OleDbType.VarChar, 255);
			OleDbParameter paramQuantite = new OleDbParameter("@QUANTITE", OleDbType.Double);
			OleDbParameter paramTempsNorm = new OleDbParameter("@TEMPS_NORM", OleDbType.Double);
			OleDbParameter paramCoutNormCons = new OleDbParameter("@COUT_NORM_CONS", OleDbType.Double);
			OleDbParameter paramCoutNormPieces = new OleDbParameter("@COUT_NORM_PIECES", OleDbType.Double);
			OleDbParameter paramCoutNormST = new OleDbParameter("@COUT_NORM_ST", OleDbType.Double);
			OleDbParameter paramNoperations = new OleDbParameter("@N_OPERATIONS", OleDbType.Double);
			OleDbParameter paramNvisites = new OleDbParameter("@N_VISITES", OleDbType.Double);
			OleDbParameter paramTempsTotal = new OleDbParameter("@TEMPS_TOTAL", OleDbType.Double);
			OleDbParameter paramCoutConsTotal = new OleDbParameter("@COUT_CONS_TOTAL", OleDbType.Double);
			OleDbParameter paramCoutPiecesTotal = new OleDbParameter("@COUT_PIECES_TOTAL", OleDbType.Double);
			OleDbParameter paramCoutSTTotal = new OleDbParameter("@COUT_ST_TOTAL", OleDbType.Double);

			OleDbConnection connection1 = new OleDbConnection(chaineConnectionProjet);
			connection1.Open();
			String strEnreg = string.Format("INSERT INTO {0}(NUM, REPERE_EQUIPEMENT, TYPE_EQUIPEMENT, QUANTITE, TEMPS_NORM, COUT_NORM_CONS, COUT_NORM_PIECES, COUT_NORM_ST, N_OPERATIONS, N_VISITES, TEMPS_TOTAL, COUT_CONS_TOTAL, COUT_PIECES_TOTAL, COUT_ST_TOTAL) VALUES({1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14})", nomTable,
				paramNum.ParameterName, paramRepere.ParameterName, paramTypeEquipement.ParameterName, paramQuantite.ParameterName,
					paramTempsNorm.ParameterName, paramCoutNormCons.ParameterName, paramCoutNormPieces.ParameterName, paramCoutNormST.ParameterName,
						paramNoperations.ParameterName, paramNvisites.ParameterName, paramTempsTotal.ParameterName,
							paramCoutConsTotal.ParameterName, paramCoutPiecesTotal.ParameterName, paramCoutSTTotal.ParameterName);

			OleDbCommand command1 = new OleDbCommand(strEnreg.ToString(), connection1);

			command1.Parameters.Add(paramNum);
			command1.Parameters.Add(paramRepere);
			command1.Parameters.Add(paramTypeEquipement);
			command1.Parameters.Add(paramQuantite);
			command1.Parameters.Add(paramTempsNorm);
			command1.Parameters.Add(paramCoutNormCons);
			command1.Parameters.Add(paramCoutNormPieces);
			command1.Parameters.Add(paramCoutNormST);
			command1.Parameters.Add(paramNoperations);
			command1.Parameters.Add(paramNvisites);
			command1.Parameters.Add(paramTempsTotal);
			command1.Parameters.Add(paramCoutConsTotal);
			command1.Parameters.Add(paramCoutPiecesTotal);
			command1.Parameters.Add(paramCoutSTTotal);

			for(int i = 0; i < dt.Rows.Count - 2; i++){ // ne pas prendre les lignes du total
				DataRow dr = dt.Rows[i];
				num = (double)(i + 1);
				repere = dr["REPERE_EQUIPEMENT"].ToString();
				typeEquipement = dr["TYPE_EQUIPEMENT"].ToString();
				quantite = double.Parse(dr["QUANTITE"].ToString());
				tempsNorm = double.Parse(dr["TEMPS_NORM"].ToString());
				coutNormCons = double.Parse(dr["COUT_NORM_CONS"].ToString());
				coutNormPieces = double.Parse(dr["COUT_NORM_PIECES"].ToString());
				coutNormST = double.Parse(dr["COUT_NORM_ST"].ToString());
				Noperations = double.Parse(dr["N_OPERATIONS"].ToString());
				Nvisites = double.Parse(dr["N_VISITES"].ToString());
				tempsTotal = double.Parse(dr["TEMPS_TOTAL"].ToString());
				coutConsTotal = double.Parse(dr["COUT_CONS_TOTAL"].ToString());
				coutPiecesTotal = double.Parse(dr["COUT_PIECES_TOTAL"].ToString());
				coutSTTotal = double.Parse(dr["COUT_ST_TOTAL"].ToString());
				paramNum.Value = num;
				paramRepere.Value = repere;
				paramTypeEquipement.Value = typeEquipement;
				paramQuantite.Value = quantite;
				paramTempsNorm.Value = tempsNorm;
				paramCoutNormCons.Value = coutNormCons;
				paramCoutNormPieces.Value = coutNormPieces;
				paramCoutNormST.Value = coutNormST;
				paramNoperations.Value = Noperations;
				paramNvisites.Value = Nvisites;
				paramTempsTotal.Value = tempsTotal;
				paramCoutConsTotal.Value = coutConsTotal;
				paramCoutPiecesTotal.Value = coutPiecesTotal;
				paramCoutSTTotal.Value = coutSTTotal;
				command1.ExecuteNonQuery();
			}
			connection1.Close();
		}

		private void creerTableOperationsMPREV_BDD(string nomTable)
		{
			string strEnreg = "CREATE TABLE " + nomTable + "(NUMERO INTEGER NOT NULL PRIMARY KEY, CODE_OPERATION VARCHAR(255), TYPE_EQUIPEMENT VARCHAR(255),"
					+ " NOM_OPERATION_MAINTENANCE VARCHAR(255), TYPE_VISITE VARCHAR(255), UNITE_USAGE VARCHAR(20), CL_MIN INTEGER,"
						+ " SYSTEMATIQUE VARCHAR(3), CONDITIONNEL VARCHAR(3),"
							+ " TMRS_MP DOUBLE PRECISION, K_FREQ DOUBLE PRECISION, REGROUPEMENT VARCHAR(3))";
			OleDbConnection connection1 = new OleDbConnection(chaineConnectionProjet);
			OleDbCommand command1 = new OleDbCommand(strEnreg);
			command1.Connection = connection1;
			connection1.Open();
			command1.ExecuteNonQuery();
			connection1.Close();
		}

		public void enregistrerTableOperationsMPREV_BDD()
		{
			string nomTable = "TABLE_LISTE_OPERATIONS_MPREV";
			// creerTableOperationsMPREV_BDD(nomTable);

			int num, classeMin;
			double tmrsMP, kFreq;
			string code, typeEquipement, nomOperation,typeVisite, uniteUsage, systematique, conditionnel, regroupement;

			OleDbParameter paramNum = new OleDbParameter("@NUM", OleDbType.Double);
			OleDbParameter paramCode = new OleDbParameter("@CODE_OPERATION", OleDbType.VarChar, 255);
			OleDbParameter paramTypeEquipement = new OleDbParameter("@TYPE_EQUIPEMENT", OleDbType.VarChar, 255);
			OleDbParameter paramNomOperation = new OleDbParameter("@NOM_OPERATION_MAINTENANCE", OleDbType.VarChar, 255);
			OleDbParameter paramTypeVisite = new OleDbParameter("@TYPE_VISITE", OleDbType.VarChar, 255);
			OleDbParameter paramUnite = new OleDbParameter("@UNITE_USAGE", OleDbType.VarChar, 20);
			OleDbParameter paramClasseMin = new OleDbParameter("@CL_MIN", OleDbType.Double);
			OleDbParameter paramSystematique = new OleDbParameter("@SYSTEMATIQUE", OleDbType.VarChar, 20);
			OleDbParameter paramConditionnel = new OleDbParameter("@CONDITIONNEL", OleDbType.VarChar, 20);
			OleDbParameter paramTmrsMP = new OleDbParameter("@TMRS_MP", OleDbType.Double);
			OleDbParameter paramKFreq = new OleDbParameter("@K_FREQ", OleDbType.Double);
			OleDbParameter paramRegroupement = new OleDbParameter("@REGROUPEMENT", OleDbType.VarChar, 3);

			OleDbConnection connection1 = new OleDbConnection(chaineConnectionProjet);
			connection1.Open();
			String strEnreg = string.Format("INSERT INTO {0}(NUMERO, CODE_OPERATION, TYPE_EQUIPEMENT, NOM_OPERATION_MAINTENANCE, TYPE_VISITE, UNITE_USAGE, CL_MIN, SYSTEMATIQUE, CONDITIONNEL, TMRS_MP, K_FREQ, REGROUPEMENT) VALUES({1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12})", nomTable,
				paramNum.ParameterName, paramCode.ParameterName, paramTypeEquipement.ParameterName, paramNomOperation.ParameterName, paramTypeVisite.ParameterName, paramUnite.ParameterName,
					paramClasseMin.ParameterName, paramSystematique.ParameterName, paramConditionnel.ParameterName,
						paramTmrsMP.ParameterName, paramKFreq.ParameterName, paramRegroupement.ParameterName);

			OleDbCommand command1 = new OleDbCommand(strEnreg.ToString(), connection1);
			command1.Parameters.Add(paramNum);
			command1.Parameters.Add(paramCode);
			command1.Parameters.Add(paramTypeEquipement);
			command1.Parameters.Add(paramNomOperation);
			command1.Parameters.Add(paramTypeVisite);
			command1.Parameters.Add(paramUnite);
			command1.Parameters.Add(paramClasseMin);
			command1.Parameters.Add(paramSystematique);
			command1.Parameters.Add(paramConditionnel);
			command1.Parameters.Add(paramTmrsMP);
			command1.Parameters.Add(paramKFreq);
			command1.Parameters.Add(paramRegroupement);

			foreach(DataRow r in listeOperationsMPREV()){

				num = int.Parse(r["NUMERO"].ToString());
				code = r["CODE_OPERATION"].ToString();
				typeEquipement = r["TYPE_EQUIPEMENT"].ToString();
				nomOperation = r["NOM_OPERATION_MAINTENANCE"].ToString();
				typeVisite = r["TYPE_VISITE"].ToString();
				uniteUsage = r["UNITE_USAGE"].ToString();
				classeMin = int.Parse(r["CL_MIN"].ToString());
				systematique = r["SYSTEMATIQUE"].ToString();
				conditionnel = r["CONDITIONNEL"].ToString();
				tmrsMP = double.Parse(r["TMRS_MP"].ToString());
				kFreq =  double.Parse(r["K_FREQ"].ToString());
				regroupement = r["REGROUPEMENT"].ToString();

				paramNum.Value = num;
				paramCode.Value = code;
				paramTypeEquipement.Value = typeEquipement;
				paramNomOperation.Value = nomOperation;
				paramTypeVisite.Value = typeVisite;
				paramUnite.Value = uniteUsage;
				paramClasseMin.Value = classeMin;
				paramSystematique.Value = systematique;
				paramConditionnel.Value = conditionnel;
				paramTmrsMP.Value = tmrsMP;
				paramKFreq.Value = kFreq;
				paramRegroupement.Value = regroupement;

				command1.ExecuteNonQuery();
			}
			connection1.Close();
		}

		public void enregistrerTableGammesOperation_BDD()
		{
			string nomTable = "TABLE_LISTE_GAMMES_OPERATION";

			OleDbParameter paramNum = new OleDbParameter("@NUM", OleDbType.Double);
			OleDbParameter paramIndice = new OleDbParameter("@INDICE_GAMME", OleDbType.Double);
			OleDbParameter paramCode = new OleDbParameter("@CODE_OPERATION", OleDbType.VarChar, 255);
			OleDbParameter paramTypeEquipement = new OleDbParameter("@TYPE_EQUIPEMENT", OleDbType.VarChar, 255);
			OleDbParameter paramConsigne = new OleDbParameter("@CONSIGNE", OleDbType.VarChar, 10);
			OleDbParameter paramNonConsigne = new OleDbParameter("@NON_CONSIGNE", OleDbType.VarChar, 10);
			OleDbParameter paramOperation1 = new OleDbParameter("@OPERATION_1", OleDbType.VarChar, 255);
			OleDbParameter paramOperation2 = new OleDbParameter("@OPERATION_2", OleDbType.VarChar, 255);
			OleDbParameter paramOperation3 = new OleDbParameter("@OPERATION_3", OleDbType.VarChar, 255);
			OleDbParameter paramOperation4 = new OleDbParameter("@OPERATION_4", OleDbType.VarChar, 255);
			OleDbParameter paramOperation5 = new OleDbParameter("@OPERATION_5", OleDbType.VarChar, 255);
			OleDbParameter paramOperation6 = new OleDbParameter("@OPERATION_6", OleDbType.VarChar, 255);
			OleDbParameter paramOperation7 = new OleDbParameter("@OPERATION_7", OleDbType.VarChar, 255);
			OleDbParameter paramOperation8 = new OleDbParameter("@OPERATION_8", OleDbType.VarChar, 255);
			OleDbParameter paramOperation9 = new OleDbParameter("@OPERATION_9", OleDbType.VarChar, 255);
			OleDbParameter paramOperation10 = new OleDbParameter("@OPERATION_10", OleDbType.VarChar, 255);
			OleDbParameter paramTempsNorm = new OleDbParameter("@TEMPS_NORM", OleDbType.Double);
			OleDbParameter paramCoutNormCons = new OleDbParameter("@COUT_NORM_CONS", OleDbType.Double);
			OleDbParameter paramCoutNormPieces = new OleDbParameter("@COUT_NORM_PIECES", OleDbType.Double);
			OleDbParameter paramCoutNormST = new OleDbParameter("@COUT_NORM_ST", OleDbType.Double);
			OleDbParameter paramListeOutillage = new OleDbParameter("@LISTE_OUTILLAGE", OleDbType.VarChar, 255);
			OleDbParameter paramListeConsommables = new OleDbParameter("@LISTE_CONSOMMABLES", OleDbType.VarChar, 255);
			OleDbParameter paramLienDocumentation = new OleDbParameter("@Lien_Documentation", OleDbType.VarChar, 255);
			OleDbParameter paramListeAutresDocuments = new OleDbParameter("@LISTE_autres_documents", OleDbType.VarChar, 255);
			OleDbParameter paramHabilitations = new OleDbParameter("@HABILITATIONS", OleDbType.VarChar, 255);
			OleDbParameter paramQualifications = new OleDbParameter("@QUALIFICATIONS", OleDbType.VarChar, 255);
			OleDbParameter paramListeCompetencesParticulieres = new OleDbParameter("@LISTE_COMPETENCES_PARTICULIERES", OleDbType.VarChar, 255);

			OleDbConnection connection1 = new OleDbConnection(chaineConnectionProjet);
			connection1.Open();
			String strEnreg = string.Format("INSERT INTO {0}(NUMERO, INDICE_GAMME, CODE_OPERATION, TYPE_EQUIPEMENT, CONSIGNE, NON_CONSIGNE, OPERATION_1, OPERATION_2, OPERATION_3, OPERATION_4, OPERATION_5," +
				"OPERATION_6, OPERATION_7, OPERATION_8, OPERATION_9, OPERATION_10,TEMPS_NORM, COUT_NORM_CONS, COUT_NORM_PIECES, COUT_NORM_ST, LISTE_OUTILLAGE, LISTE_CONSOMMABLES, Lien_Documentation, LISTE_autres_documents, HABILITATIONS, QUALIFICATIONS, LISTE_COMPETENCES_PARTICULIERES) VALUES({1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24},{25},{26},{27})",
					nomTable, paramNum.ParameterName, paramIndice.ParameterName, paramCode.ParameterName, paramTypeEquipement.ParameterName, paramConsigne.ParameterName, paramNonConsigne.ParameterName,
						paramOperation1.ParameterName, paramOperation2.ParameterName, paramOperation3.ParameterName, paramOperation4.ParameterName, paramOperation5.ParameterName,
							paramOperation6.ParameterName, paramOperation7.ParameterName, paramOperation8.ParameterName, paramOperation9.ParameterName, paramOperation10.ParameterName,
								paramTempsNorm.ParameterName, paramCoutNormCons.ParameterName, paramCoutNormPieces.ParameterName, paramCoutNormST.ParameterName,
									paramListeOutillage.ParameterName, paramListeConsommables.ParameterName, paramLienDocumentation.ParameterName, paramListeAutresDocuments.ParameterName, paramHabilitations.ParameterName, paramQualifications.ParameterName, paramListeCompetencesParticulieres.ParameterName);

			OleDbCommand command1 = new OleDbCommand(strEnreg.ToString(), connection1);
			command1.Parameters.Add(paramNum);
			command1.Parameters.Add(paramIndice);
			command1.Parameters.Add(paramCode);
			command1.Parameters.Add(paramTypeEquipement);
			command1.Parameters.Add(paramConsigne);
			command1.Parameters.Add(paramNonConsigne);
			command1.Parameters.Add(paramOperation1);
			command1.Parameters.Add(paramOperation2);
			command1.Parameters.Add(paramOperation3);
			command1.Parameters.Add(paramOperation4);
			command1.Parameters.Add(paramOperation5);
			command1.Parameters.Add(paramOperation6);
			command1.Parameters.Add(paramOperation7);
			command1.Parameters.Add(paramOperation8);
			command1.Parameters.Add(paramOperation9);
			command1.Parameters.Add(paramOperation10);
			command1.Parameters.Add(paramTempsNorm);
			command1.Parameters.Add(paramCoutNormCons);
			command1.Parameters.Add(paramCoutNormPieces);
			command1.Parameters.Add(paramCoutNormST);
			command1.Parameters.Add(paramListeOutillage);
			command1.Parameters.Add(paramListeConsommables);
			command1.Parameters.Add(paramLienDocumentation);
			command1.Parameters.Add(paramListeAutresDocuments);
			command1.Parameters.Add(paramHabilitations);
			command1.Parameters.Add(paramQualifications);
			command1.Parameters.Add(paramListeCompetencesParticulieres);

			foreach(DataRow r in listeGammesOperation()){
				if(r != null){
					paramNum.Value = int.Parse(r["NUMERO"].ToString());
					paramIndice.Value = double.Parse(r["INDICE_GAMME"].ToString());
					paramCode.Value = r["CODE_OPERATION"].ToString();
					paramTypeEquipement.Value = r["TYPE_EQUIPEMENT"].ToString();
					paramConsigne.Value = r["CONSIGNE"].ToString();
					paramNonConsigne.Value = r["NON_CONSIGNE"].ToString();
					paramOperation1.Value = r["OPERATION_1"].ToString();
					paramOperation2.Value = r["OPERATION_2"].ToString();
					paramOperation3.Value = r["OPERATION_3"].ToString();
					paramOperation4.Value = r["OPERATION_4"].ToString();
					paramOperation5.Value = r["OPERATION_5"].ToString();
					paramOperation6.Value = r["OPERATION_6"].ToString();
					paramOperation7.Value = r["OPERATION_7"].ToString();
					paramOperation8.Value = r["OPERATION_8"].ToString();
					paramOperation9.Value = r["OPERATION_9"].ToString();
					paramOperation10.Value = r["OPERATION_10"].ToString();
					paramTempsNorm.Value = r["TEMPS_NORM"].ToString();
					paramCoutNormCons.Value = r["COUT_NORM_CONS"].ToString();
					paramCoutNormPieces.Value = r["COUT_NORM_PIECES"].ToString();
					paramCoutNormST.Value = r["COUT_NORM_ST"].ToString();
					paramListeOutillage.Value = r["LISTE_OUTILLAGE"].ToString();
					paramListeConsommables.Value = r["LISTE_CONSOMMABLES"].ToString();
					paramLienDocumentation.Value = r["Lien_Documentation"].ToString();
					paramListeAutresDocuments.Value = r["LISTE_autres_documents"].ToString();
					paramHabilitations.Value = r["HABILITATIONS"].ToString();
					paramQualifications.Value = r["QUALIFICATIONS"].ToString();
					paramListeCompetencesParticulieres.Value = r["LISTE_COMPETENCES_PARTICULIERES"].ToString();

					command1.ExecuteNonQuery();
				}
			}
			connection1.Close();
		}

		public void enregistrerTableTauxHorairesBDD()
		{
			DataTable dt = dsProjet.Tables["Table TAUX_HORAIRES"];
			// double num = 0;
			string qualification, source, tauxBase;
			string nomTable = "TABLE_LISTE_TAUX_HORAIRES";

			// OleDbParameter paramNum = new OleDbParameter("@NUM2", OleDbType.Double);
			OleDbParameter paramQualification = new OleDbParameter("@QUALIFICATION", OleDbType.VarChar, 50);
			OleDbParameter paramSource = new OleDbParameter("@SOURCE", OleDbType.VarChar, 50);
			OleDbParameter paramTauxBase = new OleDbParameter("@TAUX_BASE", OleDbType.Double);

			OleDbConnection connection1 = new OleDbConnection(chaineConnectionProjet);
			connection1.Open();
			String strEnreg = string.Format("INSERT INTO {0}(QUALIFICATION, SOURCE, TAUX_BASE) VALUES({1},{2},{3})", nomTable, paramQualification.ParameterName, paramSource.ParameterName, paramTauxBase.ParameterName);
			OleDbCommand command1 = new OleDbCommand(strEnreg.ToString(), connection1);

			command1.Parameters.Add(paramQualification);
			command1.Parameters.Add(paramSource);
			command1.Parameters.Add(paramTauxBase);

			for(int i = 0; i < dt.Rows.Count; i++){

				DataRow dr = dt.Rows[i];
				// num = (double)(i + 1);
				qualification = dr["QUALIFICATION"].ToString();
				source = dr["SOURCE"].ToString();
				tauxBase = dr["TAUX_BASE"].ToString();
				paramQualification.Value = qualification;
				paramSource.Value = source;
				paramTauxBase.Value = tauxBase;

				command1.ExecuteNonQuery();
			}
			connection1.Close();
		}

		private string chaineLocalisation(string secteur, string batiment, string emplacement)
		{
			string localisation = "";
			if(secteur =="" & batiment == "" & emplacement ==""){
				localisation = "";
			}else{
				if(batiment == "" & emplacement == ""){
						localisation = secteur;
				}else{
					if(emplacement == ""){
					localisation = secteur + " / " + batiment;
				}else{
						localisation = secteur + " / " + batiment + " / " + emplacement;
					}
				}
			}
			return localisation;
		}

		private DataRow[] listeOperationsMPREV()
		{
			DataTable dt = afficheDsProjet.Tables["Table Plan de maintenance [par REPERE_EQUIPEMENT]"];
			DataTable dt1 = selectDistinct("", dt, "CODE_OPERATION");
			DataRow[] tabdr = new DataRow[dt1.Rows.Count];
			DataTable dt2 = detailOperationsMPREV();
			for(int i = 0; i < tabdr.Length; i++){
				string code = dt1.Rows[i]["CODE_OPERATION"].ToString();
				string selection = "CODE_OPERATION = '" + code + "'";
				DataRow[] dr1 = dt2.Select(selection);
				tabdr[i] = dr1[0];
			}
			return  tabdr;
		}

		private DataRow[] listeGammesOperation()
		{
			DataTable dt = afficheDsProjet.Tables["Table Plan de maintenance [par REPERE_EQUIPEMENT]"];
			DataTable dt1 = selectDistinct("", dt, "CODE_OPERATION");
			DataRow[] tabdr = new DataRow[dt1.Rows.Count];
			DataTable dt2 = detailGammesOperation();
			for(int i = 0; i < tabdr.Length; i++){
				string code = dt1.Rows[i]["CODE_OPERATION"].ToString();
				string selection = "CODE_OPERATION = '" + code + "'";
				DataRow[] dr1 = dt2.Select(selection);
				if(dr1.Length > 0){
				// MessageBox.Show(i.ToString());
				tabdr[i] = dr1[dr1.Length-1];
				}
				else{
					tabdr[i] = null;
				}
			}
			return  tabdr;
		}

		private DataTable detailOperationsMPREV()
		{
			DataTable dt;
			string chaineConnection = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + "\\" + nomBaseOperations + ";Jet OLEDB:Database Password=mexico";
			OleDbConnection connection2 = new OleDbConnection(chaineConnection);
			connection2.Open();
			string strSQLOperations = "SELECT NUMERO,CODE_OPERATION,TYPE_EQUIPEMENT,NOM_OPERATION_MAINTENANCE,TYPE_VISITE,UNITE_USAGE,CL_MIN,SYSTEMATIQUE,CONDITIONNEL,TMRS_MP,K_FREQ,REGROUPEMENT FROM Liste_OPERATIONS_MPREV";
			OleDbDataAdapter oleDbAdapter2 = new OleDbDataAdapter (strSQLOperations, connection2);
			DataSet dsOperationsMPREV = new DataSet();
			oleDbAdapter2.Fill(dsOperationsMPREV);
			connection2.Close();
			dt = dsOperationsMPREV.Tables[0];
			return dt;
		}

		private DataTable detailGammesOperation()
		{
			DataTable dt;
			string chaineConnection = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + "\\" + nomBaseOperations + ";Jet OLEDB:Database Password=mexico";
			OleDbConnection connection2 = new OleDbConnection(chaineConnection);
			connection2.Open();
			string strSQLOperations = "SELECT NUMERO,INDICE_GAMME,CODE_OPERATION,TYPE_EQUIPEMENT,CONSIGNE,NON_CONSIGNE,OPERATION_1,OPERATION_2,OPERATION_3,OPERATION_4,OPERATION_5," +
				"OPERATION_6,OPERATION_7,OPERATION_8,OPERATION_9,OPERATION_10," +
					"TEMPS_NORM,COUT_NORM_CONS,COUT_NORM_PIECES,COUT_NORM_ST," +
						"LISTE_OUTILLAGE,LISTE_CONSOMMABLES,Lien_Documentation,LISTE_autres_documents,HABILITATIONS,QUALIFICATIONS,LISTE_COMPETENCES_PARTICULIERES FROM Liste_GAMMES_OPERATION";
			OleDbDataAdapter oleDbAdapter2 = new OleDbDataAdapter (strSQLOperations, connection2);
			DataSet dsGammesOperation = new DataSet();
			oleDbAdapter2.Fill(dsGammesOperation);
			connection2.Close();
			dt = dsGammesOperation.Tables[0];
			return dt;
		}

		public DataTable creerTableListeCables()
		{
			DataTable dt = new DataTable("Table Liste CABLES");

			// données venant de tableauCaracteristiques[120]
			dt.Columns.Add("DESIGNATION_CIRCUIT", System.Type.GetType("System.String"));
			dt.Columns.Add("REPERE_CIRCUIT", System.Type.GetType("System.String"));
			dt.Columns.Add("PCC_RESEAU_AMONT", System.Type.GetType("System.String"));
			dt.Columns.Add("RSURX_RESEAU_AMONT", System.Type.GetType("System.String"));
			dt.Columns.Add("ICC_AMONT_EN_BT", System.Type.GetType("System.String"));
			dt.Columns.Add("AUTORISATION_TOLERANCE_5_POUR_CENT", System.Type.GetType("System.String"));
			dt.Columns.Add("TABLEAU_SECTIONS_NORMALISEES", System.Type.GetType("System.String"));
			dt.Columns.Add("N_MAXI_CONDUCTEURS_EN_PARALLELE", System.Type.GetType("System.String"));
			dt.Columns.Add("NATURE_CONDUCTEUR", System.Type.GetType("System.String"));
			dt.Columns.Add("TYPE_ISOLANT", System.Type.GetType("System.String"));
			dt.Columns.Add("TYPE_CIRCUIT", System.Type.GetType("System.String"));
			dt.Columns.Add("TENSION_NOMINALE", System.Type.GetType("System.String"));
			dt.Columns.Add("COURANT_EMPLOI", System.Type.GetType("System.String"));
			dt.Columns.Add("LONGUEUR", System.Type.GetType("System.String"));
			dt.Columns.Add("CHUTE_TENSION_ADMISSIBLE_EN_POUR_CENT", System.Type.GetType("System.String"));
			dt.Columns.Add("MODE_POSE", System.Type.GetType("System.String"));
			dt.Columns.Add("CODE_POSE", System.Type.GetType("System.String"));
			dt.Columns.Add("METHODE_REFERENCE", System.Type.GetType("System.String"));
			dt.Columns.Add("CODE_METHODE", System.Type.GetType("System.String"));
			dt.Columns.Add("TEMPERATURE_AMBIANTE", System.Type.GetType("System.String"));
			dt.Columns.Add("TEMPERATURE_SOL", System.Type.GetType("System.String"));
			dt.Columns.Add("NATURE_TERRAIN", System.Type.GetType("System.String"));
			dt.Columns.Add("RESISTIVITE_THERMIQUE_SOL", System.Type.GetType("System.String"));
			dt.Columns.Add("N_CIRCUITS_OU_CABLES_JOINTIFS", System.Type.GetType("System.String"));
			dt.Columns.Add("N_COUCHES", System.Type.GetType("System.String"));
			dt.Columns.Add("N_CONDUITS_EN_SOL", System.Type.GetType("System.String"));
			dt.Columns.Add("DISTANCE_ENTRE_CONDUITS", System.Type.GetType("System.String"));
			dt.Columns.Add("N_CABLES_OU_CIRCUITS_EN_SOL", System.Type.GetType("System.String"));
			dt.Columns.Add("DISTANCE_ENTRE_CABLES", System.Type.GetType("System.String"));
			dt.Columns.Add("PRESENCE_RISQUE_BE3", System.Type.GetType("System.String"));
			dt.Columns.Add("NEUTRE_CHARGE_15", System.Type.GetType("System.String"));
			dt.Columns.Add("POSE_A_2_DIAMETRES", System.Type.GetType("System.String"));
			dt.Columns.Add("POSE_NON_SYMETRIQUE", System.Type.GetType("System.String"));
			dt.Columns.Add("MODE_POSE_AVEC_CODE", System.Type.GetType("System.String"));
			dt.Columns.Add("TYPE_PROTECTION", System.Type.GetType("System.String"));
			dt.Columns.Add("TYPE", System.Type.GetType("System.String"));
			dt.Columns.Add("CALIBRE", System.Type.GetType("System.String"));
			dt.Columns.Add("N_POL_DECL", System.Type.GetType("System.String"));
			dt.Columns.Add("TYPE_DECLENCHEUR", System.Type.GetType("System.String"));
			dt.Columns.Add("IR", System.Type.GetType("System.String"));
			dt.Columns.Add("IM", System.Type.GetType("System.String"));
			dt.Columns.Add("k_IR", System.Type.GetType("System.String"));
			dt.Columns.Add("k_IM", System.Type.GetType("System.String"));
			dt.Columns.Add("EN_ATTENTE_2", System.Type.GetType("System.String"));
			dt.Columns.Add("EN_ATTENTE_3", System.Type.GetType("System.String"));
			dt.Columns.Add("COURBE_DISJONCTEUR_DOMESTIQUE", System.Type.GetType("System.String"));
			dt.Columns.Add("REGLAGE_DECLENCHEUR_THERMIQUE", System.Type.GetType("System.String"));
			dt.Columns.Add("REGLAGE_DECLENCHEUR_MAGNETIQUE", System.Type.GetType("System.String"));
			dt.Columns.Add("EN_ATTENTE_4", System.Type.GetType("System.String"));
			dt.Columns.Add("CALIBRE_FUSIBLE_gG", System.Type.GetType("System.String"));
			dt.Columns.Add("EN_ATTENTE_5", System.Type.GetType("System.String"));
			dt.Columns.Add("EN_ATTENTE_6", System.Type.GetType("System.String"));
			dt.Columns.Add("EN_ATTENTE_7", System.Type.GetType("System.String"));
			dt.Columns.Add("EN_ATTENTE_8", System.Type.GetType("System.String"));
			dt.Columns.Add("EN_ATTENTE_9", System.Type.GetType("System.String"));
			dt.Columns.Add("SCHEMA_LIAISONS_TERRE", System.Type.GetType("System.String"));
			dt.Columns.Add("SECTION_MAXI_EN_CABLES_MULTICONDUCTEURS", System.Type.GetType("System.String"));
			dt.Columns.Add("NEUTRE_CHARGE_33", System.Type.GetType("System.String"));
			dt.Columns.Add("FACTEUR_SURDIMENSIONNEMENT", System.Type.GetType("System.String"));
			dt.Columns.Add("NOM_PROJET", System.Type.GetType("System.String"));
			dt.Columns.Add("IN", System.Type.GetType("System.String"));
			dt.Columns.Add("IZ_DECLASSE", System.Type.GetType("System.String"));
			dt.Columns.Add("FACTEUR_GLOBAL_CORRECTION", System.Type.GetType("System.String"));
			dt.Columns.Add("SECTION_PHASE_VALIDEE", System.Type.GetType("System.String"));
			dt.Columns.Add("SECTION_NEUTRE_VALIDEE", System.Type.GetType("System.String"));
			dt.Columns.Add("SECTION_PE_VALIDEE", System.Type.GetType("System.String"));
			dt.Columns.Add("DELTA_U_VOLTS_SECTION_VALIDEE", System.Type.GetType("System.String"));
			dt.Columns.Add("ICC_AVAL_SECTION_VALIDEE", System.Type.GetType("System.String"));
			//
			//
			dt.Columns.Add("REPERE_TENANT", System.Type.GetType("System.String"));
			dt.Columns.Add("REPERE_ABOUTISSANT", System.Type.GetType("System.String"));
			//
			//
			dt.Columns.Add("DATE_CALCUL", System.Type.GetType("System.String"));
			dt.Columns.Add("AUTEUR_CALCUL", System.Type.GetType("System.String"));
			dt.Columns.Add("INDICE_REVISION", System.Type.GetType("System.String"));
			//
			//
			dt.Columns.Add("DEMARRAGE_MOTEUR", System.Type.GetType("System.String"));
			dt.Columns.Add("TYPE_DEMARRAGE", System.Type.GetType("System.String"));
			dt.Columns.Add("ID_SUR_IN", System.Type.GetType("System.String"));
			dt.Columns.Add("IN_MOTEUR", System.Type.GetType("System.String"));
			//
			//
			dt.Columns.Add("K_METHODE", System.Type.GetType("System.String"));
			dt.Columns.Add("N_CONDUITS_VERTICAL", System.Type.GetType("System.String"));
			dt.Columns.Add("N_CONDUITS_HORIZONTAL", System.Type.GetType("System.String"));
			dt.Columns.Add("N_CABLES_PAR_CONDUIT", System.Type.GetType("System.String"));
			//
			dt.Columns.Add("TYPE_CABLE", System.Type.GetType("System.String"));
			//
			dt.Columns.Add("DDR_DISJ_INDUSTRIEL", System.Type.GetType("System.String"));
			dt.Columns.Add("I_DELTA_N_DISJ_INDUSTRIEL", System.Type.GetType("System.String"));
			dt.Columns.Add("RETARD_DISJ_INDUSTRIEL", System.Type.GetType("System.String"));
			//
			dt.Columns.Add("DDR_DISJ_DOMESTIQUE", System.Type.GetType("System.String"));
			dt.Columns.Add("I_DELTA_N_DISJ_DOMESTIQUE", System.Type.GetType("System.String"));
			dt.Columns.Add("RETARD_DISJ_DOMESTIQUE", System.Type.GetType("System.String"));
			//
			dt.Columns.Add("SECTION_COMPLETE_FORMATTEE", System.Type.GetType("System.String"));
			//

			// données venant de tableauSections[50]
			dt.Columns.Add("SECTION_PAR_PHASE_VALIDEE", System.Type.GetType("System.String"));
			dt.Columns.Add("N_COND_PAR_PHASE_SECTION_VALIDEE", System.Type.GetType("System.String"));

			return dt;
		}

		public void ajouterCableBT(cableBT cable)
		{
			DataTable dt = dsProjet.Tables["Table Liste CABLES"];
			DataRow dr = dt.NewRow();
			configurerLigne(dr, cable);
			dt.Rows.Add(dr);
		}

		public void configurerLigne(DataRow dr, cableBT cable)
		{
			dr["DESIGNATION_CIRCUIT"] = cable.afficheCar(0);
			dr["REPERE_CIRCUIT"] = cable.afficheCar(1);
			dr["PCC_RESEAU_AMONT"] = cable.afficheCar(2);
			dr["RSURX_RESEAU_AMONT"] = cable.afficheCar(3);
			dr["ICC_AMONT_EN_BT"] = cable.afficheCar(4);
			dr["AUTORISATION_TOLERANCE_5_POUR_CENT"] = cable.afficheCar(5);
			dr["TABLEAU_SECTIONS_NORMALISEES"] = cable.afficheCar(6);
			dr["N_MAXI_CONDUCTEURS_EN_PARALLELE"] = cable.afficheCar(7);
			dr["NATURE_CONDUCTEUR"] = cable.afficheCar(8);
			dr["TYPE_ISOLANT"] = cable.afficheCar(9);
			dr["TYPE_CIRCUIT"] = cable.afficheCar(10);
			dr["TENSION_NOMINALE"] = cable.afficheCar(11);
			dr["COURANT_EMPLOI"] = cable.afficheCar(12);
			dr["LONGUEUR"] = cable.afficheCar(13);
			dr["CHUTE_TENSION_ADMISSIBLE_EN_POUR_CENT"] = cable.afficheCar(14);
			dr["MODE_POSE"] = cable.afficheCar(15);
			dr["CODE_POSE"] = cable.afficheCar(16);
			dr["METHODE_REFERENCE"] = cable.afficheCar(17);
			dr["CODE_METHODE"] = cable.afficheCar(18);
			dr["TEMPERATURE_AMBIANTE"] = cable.afficheCar(19);
			dr["TEMPERATURE_SOL"] = cable.afficheCar(20);
			dr["NATURE_TERRAIN"] = cable.afficheCar(21);
			dr["RESISTIVITE_THERMIQUE_SOL"] = cable.afficheCar(22);
			dr["N_CIRCUITS_OU_CABLES_JOINTIFS"] = cable.afficheCar(23);
			dr["N_COUCHES"] = cable.afficheCar(24);
			dr["N_CONDUITS_EN_SOL"] = cable.afficheCar(25);
			dr["DISTANCE_ENTRE_CONDUITS"] = cable.afficheCar(26);
			dr["N_CABLES_OU_CIRCUITS_EN_SOL"] = cable.afficheCar(27);
			dr["DISTANCE_ENTRE_CABLES"] = cable.afficheCar(28);
			dr["PRESENCE_RISQUE_BE3"] = cable.afficheCar(29);
			dr["NEUTRE_CHARGE_15"] = cable.afficheCar(30);
			dr["POSE_A_2_DIAMETRES"] = cable.afficheCar(31);
			dr["POSE_NON_SYMETRIQUE"] = cable.afficheCar(32);
			dr["MODE_POSE_AVEC_CODE"] = cable.afficheCar(33);
			dr["TYPE_PROTECTION"] = cable.afficheCar(34);
			dr["TYPE"] = cable.afficheCar(35);
			dr["CALIBRE"] = cable.afficheCar(36);
			dr["N_POL_DECL"] = cable.afficheCar(37);
			dr["TYPE_DECLENCHEUR"] = cable.afficheCar(38);
			dr["IR"] = cable.afficheCar(39);
			dr["IM"] = cable.afficheCar(40);
			dr["k_IR"] = cable.afficheCar(41);
			dr["k_IM"] = cable.afficheCar(42);
			dr["EN_ATTENTE_2"] = "";
			dr["EN_ATTENTE_3"] = "";
			dr["COURBE_DISJONCTEUR_DOMESTIQUE"] = cable.afficheCar(45);
			dr["REGLAGE_DECLENCHEUR_THERMIQUE"] = cable.afficheCar(46);
			dr["REGLAGE_DECLENCHEUR_MAGNETIQUE"] = cable.afficheCar(47);
			dr["EN_ATTENTE_4"] = "";
			dr["CALIBRE_FUSIBLE_gG"] = cable.afficheCar(49);
			dr["EN_ATTENTE_5"] = "";
			dr["EN_ATTENTE_6"] = "";
			dr["EN_ATTENTE_7"] = "";
			dr["EN_ATTENTE_8"] = "";
			dr["EN_ATTENTE_9"] = "";
			dr["SCHEMA_LIAISONS_TERRE"] = cable.afficheCar(55);
			dr["SECTION_MAXI_EN_CABLES_MULTICONDUCTEURS"] = cable.afficheCar(56);
			dr["NEUTRE_CHARGE_33"] = cable.afficheCar(57);
			dr["FACTEUR_SURDIMENSIONNEMENT"] = cable.afficheCar(58);
			dr["NOM_PROJET"] = cable.afficheCar(59);
			//
			//
			dr["REPERE_TENANT"] = cable.afficheCar(70);
			dr["REPERE_ABOUTISSANT"] = cable.afficheCar(71);
			//
			dr["DATE_CALCUL"] = cable.afficheCar(75);
			dr["AUTEUR_CALCUL"] = cable.afficheCar(76);
			dr["INDICE_REVISION"] = cable.afficheCar(77);
			//
			dr["DEMARRAGE_MOTEUR"] = cable.afficheCar(80);
			dr["TYPE_DEMARRAGE"] = cable.afficheCar(81);
			dr["ID_SUR_IN"] = cable.afficheCar(82);
			dr["IN_MOTEUR"] = cable.afficheCar(83);
			//
			dr["K_METHODE"] = cable.afficheCar(90);
			dr["N_CONDUITS_VERTICAL"] = cable.afficheCar(91);
			dr["N_CONDUITS_HORIZONTAL"] = cable.afficheCar(92);
			dr["N_CABLES_PAR_CONDUIT"] = cable.afficheCar(93);
			//
			dr["TYPE_CABLE"] = cable.afficheCar(95);
			//
			dr["DDR_DISJ_INDUSTRIEL"] = cable.afficheCar(100);
			dr["I_DELTA_N_DISJ_INDUSTRIEL"] = cable.afficheCar(101);
			dr["RETARD_DISJ_INDUSTRIEL"] = cable.afficheCar(102);
			//
			dr["DDR_DISJ_DOMESTIQUE"] = cable.afficheCar(105);
			dr["I_DELTA_N_DISJ_DOMESTIQUE"] = cable.afficheCar(106);
			dr["RETARD_DISJ_DOMESTIQUE"] = cable.afficheCar(107);

			// configurer les résultats
			dr["IN"] = cable.afficheRes(0).ToString("");
			dr["IZ_DECLASSE"] = cable.afficheRes(1).ToString("f1");

			// configurer les coefficients
			dr["FACTEUR_GLOBAL_CORRECTION"] = cable.afficheCoeff(3).ToString("f4");

			// configurer les sections
			if(cable.afficheRes(34) == 1){ // Test si section en câbles monoconducteurs
				dr["SECTION_PHASE_VALIDEE"] = cable.afficheSection(18) + " x " + cable.afficheSection(17);
			}else{
				dr["SECTION_PHASE_VALIDEE"] = cable.afficheSection(17);
			}
//			dr["SECTION_NEUTRE_VALIDEE"] = cable.afficheCar(64);
//			dr["SECTION_PE_VALIDEE"] = cable.afficheCar(65);

			// configurer les résultats pour la section validée
			dr["DELTA_U_VOLTS_SECTION_VALIDEE"] = cable.afficheRes(22).ToString("f2");
			dr["ICC_AVAL_SECTION_VALIDEE"] = cable.afficheRes(24).ToString("f1");
		}

		public bool testExistenceRepereCable(string repere)
		{
			bool result = true;
			string strSQL = "REPERE_CIRCUIT='" + repere + "'";
			DataRow[] drTab = dsProjet.Tables[0].Select(strSQL);
			if(drTab.Length == 0){result = false;}
			return result;
		}

		public int rechercherIndexCableBT(string repere)
		{
			int index = -1;
			DataTable dt = dsProjet.Tables["Table Liste CABLES"];
			int j = 0;
			foreach(DataRow dr in dt.Rows){
				if((string)dr["REPERE_CIRCUIT"] == repere){
					index = j;
				}
				j++;
			}
			return index; // retourne -1 si non trouvé
		}
	}

	//-----------------------------------------------------------------------------------------------------

	public class visu
	{

		private string[] _systemes;
		private string[] _batiments;
		private string _classe_G;
		private bool _appartenance;
		private bool _index = true;
		private string _chaineTri;
		private string _modeAffichage;


		public visu(string[] systemes, string[] batiments, string classeG, bool appartenance)
		{
			this._systemes = systemes;
			this._batiments = batiments;
			this._classe_G = classeG;
			this._appartenance = appartenance;
			this._modeAffichage = "LISTE";
		}

		public string[] systemes{
			get {return _systemes;}
			set{_systemes = value;}
		}

		public string[] batiments{
			get {return _batiments;}
			set{_batiments = value;}
		}

		public string classe_G{
			get {return _classe_G;}
			set{_classe_G = value;}
		}

		public bool appartenance{
			get {return _appartenance;}
			set{_appartenance = value;}
		}

		public bool afficherIndex{
			get {return _index;}
			set{_index = value;}
		}

		public string chaineTri{
			get {return _chaineTri;}
			set{_chaineTri = value;}
		}

		public string modeAffichage{
			get {return _modeAffichage;}
			set{_modeAffichage = value;}
		}

		private bool testConfigSys()
		{
			bool res = false;
			foreach(string item in systemes){
				if(item != String.Empty){
					bool temp = true;
					res = res | temp;
				}
			}
			return res;
		}

		private bool testConfigBat()
		{
			bool res = false;
			if(batiments.Length != 0){
				res = true;
			}
			return res;
		}

		public bool config()
		{
			bool res = false;
			res = testConfigSys() && testConfigBat();
			return res;
		}

		public string chaineSelection()
		{
			string chaineSelection = String.Empty;
			string chaineRechercheSystemes = String.Empty;
			string chaineRechercheBatiments = String.Empty;
			string chaineRechercheClasseG = String.Empty;
			string chaineAppartenance = String.Empty;


			// configuration liste SYSTEMES
			for(int i = 0; i < systemes.Length; i++){
				if(systemes[i] != String.Empty && systemes[i] != null){
					if(chaineRechercheSystemes == ""){
						chaineRechercheSystemes = "'" + systemes[i] + "'";
					}
					else{
						chaineRechercheSystemes = chaineRechercheSystemes + " OR NOM_SYSTEME='" + systemes[i] + "'";
					}
				}
			}
			if(chaineRechercheSystemes != String.Empty){
				chaineRechercheSystemes = "(NOM_SYSTEME = " + chaineRechercheSystemes +  ")";
			}

			// configuration liste BATIMENTS
			for(int i = 0; i < batiments.Length; i++){
				if(batiments[i] != String.Empty && batiments[i] != null){
					if(chaineRechercheBatiments == ""){
						chaineRechercheBatiments = "'" + batiments[i] + "'";
					}
					else{
						chaineRechercheBatiments = chaineRechercheBatiments + " OR REPERE_BATIMENT='" + batiments[i] + "'";
					}
				}
			}
			if(chaineRechercheBatiments != String.Empty){
				chaineRechercheBatiments =  "(REPERE_BATIMENT = " + chaineRechercheBatiments + ")";
			}

			// configuration liste CLASSE_G
			chaineRechercheClasseG = _classe_G;
			if(chaineRechercheClasseG != String.Empty){
				chaineRechercheClasseG = "(CLASSE_G = '" + _classe_G + "')";
			}

			// configuration liste APPARTENANCE
			if(_appartenance == true){
				chaineAppartenance = "(APPARTENANCE_PERIMETRE = 'OUI')";
			}

			// construction chaine
			string[] tab = new string[]{chaineRechercheSystemes, chaineRechercheBatiments, chaineRechercheClasseG, chaineAppartenance};
			chaineSelection = tab[0];
			string separateur = " AND ";
			string memoire = chaineSelection;

			for(int j = 1; j < tab.Length; j++){
				if(tab[j-1] != String.Empty && tab[j] != String.Empty){
					chaineSelection = memoire + separateur + tab[j];
					tab[j] = chaineSelection;
					memoire = chaineSelection;
				}else if(tab[j] != String.Empty){
					if(memoire != String.Empty){
						chaineSelection = memoire + separateur + tab[j];
					}else{
						chaineSelection = memoire + tab[j];
					}
					tab[j] = chaineSelection;
					memoire = chaineSelection;
				}
			}
			return chaineSelection;
		}
	}

	//-----------------------------------------------------------------------------------------------------

	public class utilisateur
	{

		private string nomUtilisateur;
		private int niveauAcces;
		private string[] listeAbonnement;
		private string password; // non_encoded password

		private string nomBaseOperations = "BASE MAINTENANCE_OPERATIONS ET FMD.mdb";
		private string chaineConnection = String.Empty;

		private typEtat _etat;

		public enum typEtat { vide, enCours, configure, modifie, enregistre, enSelection };


		public utilisateur(string nom, int niveau, string[] abo, string password) {

			initialise(nom, niveau, abo, password);

		}

		private void initialise(string nom, int niveau, string[] abo, string psswd) {

			this.nomUtilisateur = nom;
			this.niveauAcces = niveau;
			this.listeAbonnement = abo;
			this.password = psswd;
			this._etat = typEtat.vide;
			this.chaineConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + "\\" + nomBaseOperations + ";Jet OLEDB:Database Password=mexico";
		}

		public string afficheNom{
			get {return nomUtilisateur;}
		}

		public int afficheNiveauAcces{
			get {return niveauAcces;}
		}

		public string[] afficheAbonnement{
			get {return listeAbonnement;}
		}

		public string afficheNomBase{
			get {return nomBaseOperations;}
		}

		public typEtat etat{
			set{_etat = value;}
			get {return _etat;}
		}

		public string affiche_enc_Psswd(){

			string res = getEncryptedPass();
			return res;
		}

		public void modifieNom(string _nom){
			this.nomUtilisateur = _nom;
		}

		public void modifieNiveauAcces(int _niv){
			this.niveauAcces = _niv;
		}

		public void modifieMotPasse(string _password){
			this.password = _password;
		}

		public void modifieAbonnement(string[] _abo){
			this.listeAbonnement = _abo;
		}

		public void modifieItemAbonnement(int n, string value){
			this.listeAbonnement[n] = value;
		}

		public void initialiseAbonnement(){
			this.listeAbonnement = lireAbonnement();
		}

		public void initialiseNiveauAcces(){
			this.niveauAcces = lireNiveauAcces();
		}

		public bool testAbonnement(int n)
		{
			bool res = false;
			switch(n){
			case 0 :
				if(listeAbonnement[0] == "1"){res = true;}else{res = false;}
			break;
			case 1 :
				if(listeAbonnement[1] == "1"){res = true;}else{res = false;}
			break;
			case 2 :
				if(listeAbonnement[2] == "1"){res = true;}else{res = false;}
			break;
			case 3 :
				if(listeAbonnement[3] == "1"){res = true;}else{res = false;}
			break;
			case 4 :
				if(listeAbonnement[4] == "1"){res = true;}else{res = false;}
			break;
			case 5 :
				if(listeAbonnement[5] == "1"){res = true;}else{res = false;}
			break;
			case 6 :
				if(listeAbonnement[6] == "1"){res = true;}else{res = false;}
			break;
			case 7 :
				if(listeAbonnement[7] == "1"){res = true;}else{res = false;}
			break;
			case 8 :
				if(listeAbonnement[8] == "1"){res = true;}else{res = false;}
			break;
			case 9 :
				if(listeAbonnement[9] == "1"){res = true;}else{res = false;}
			break;
			default :
				res = false;
			break;
			}
			return res;
		}

		private string[] lireAbonnement()
		{
			string[] tab = new string[10];
			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			string strSQL = "SELECT  LISTE_ABONNEMENT FROM TABLE_UTILISATEURS WHERE NOM_UTILISATEUR='" + nomUtilisateur + "'";
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter (strSQL, connection1);

			DataSet ds = new DataSet();
			oleDbAdapter1.Fill(ds);
			connection1.Close();
			string listeAbo = ds.Tables[0].Rows[0]["LISTE_ABONNEMENT"].ToString();
			tab = listeAbo.Split(new Char[] {' '});
			return tab;
		}

		public bool existUtilisateur()
		{
			bool res = false;
			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();

			string strSQL = "SELECT  NOM_UTILISATEUR FROM TABLE_UTILISATEURS WHERE NOM_UTILISATEUR='" + nomUtilisateur + "'";
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter (strSQL, connection1);
			DataSet ds = new DataSet();
			oleDbAdapter1.Fill(ds);
			connection1.Close();
			if(ds.Tables[0].Rows.Count != 0){res = true;}
			return res;
		}

		public int lireNiveauAcces()
		{
			int res = -1;
			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			string strSQL = "SELECT NOM_UTILISATEUR, MOT_PASSE, NIVEAU_ACCES FROM TABLE_UTILISATEURS WHERE (NOM_UTILISATEUR='" + nomUtilisateur + "')";
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter (strSQL, connection1);

			DataSet ds = new DataSet();
			oleDbAdapter1.Fill(ds);
			connection1.Close();
			res = int.Parse(ds.Tables[0].Rows[0]["NIVEAU_ACCES"].ToString());
			return res;
		}

		public bool checkMotPasse(string non_enc_psswd)
		{
			bool res = false;
			string enc_psswd = getEncryptedData(non_enc_psswd);
			if(lire_enc_motPasse() == enc_psswd) {res = true;}else{res = false;}
			return res;
		}

		public string lire_enc_motPasse()
		{
			string res = String.Empty;
			string strSQL = "SELECT NOM_UTILISATEUR, MOT_PASSE FROM TABLE_UTILISATEURS WHERE (NOM_UTILISATEUR='" + nomUtilisateur + "')";
			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter(strSQL, connection1);
			connection1.Close();
			DataSet ds = new DataSet();
			oleDbAdapter1.Fill(ds);
			res = ds.Tables[0].Rows[0]["MOT_PASSE"].ToString();
			return res;
		}

		private string convertirListeAbo(string[] tab)
		{
			string res = String.Empty;
			for(int i = 0; i < tab.Length; i++){
				if(i < tab.Length - 1){res = res + tab[i] + " ";}else{res = res + tab[i];}
			}
			return res;
		}

		public void enregistrerUtilisateur()
		{
			string nomTable = "TABLE_UTILISATEURS";
			string listeAbo = convertirListeAbo(listeAbonnement);
			string enc_motPasse = getEncryptedPass();

			// définition des paramètres
			OleDbParameter paramNom = new OleDbParameter("@NOM_UTILISATEUR", OleDbType.VarChar, 50);
			OleDbParameter paramMot = new OleDbParameter("@MOT_PASSE", OleDbType.VarChar, 50);
			OleDbParameter paramNiveau = new OleDbParameter("@NIVEAU_ACCES", OleDbType.Double);
			OleDbParameter paramAbo = new OleDbParameter("@LISTE_ABONNEMENT", OleDbType.VarChar, 50);

			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			String strEnreg = string.Format("INSERT INTO {0}(NOM_UTILISATEUR, MOT_PASSE, NIVEAU_ACCES, LISTE_ABONNEMENT) VALUES({1},{2},{3},{4})", nomTable, paramNom.ParameterName, paramMot.ParameterName, paramNiveau.ParameterName, paramAbo.ParameterName);
			OleDbCommand command1 = new OleDbCommand(strEnreg.ToString(), connection1);

			command1.Parameters.Add(paramNom);
			command1.Parameters.Add(paramMot);
			command1.Parameters.Add(paramNiveau);
			command1.Parameters.Add(paramAbo);

			paramNom.Value = nomUtilisateur;
			paramMot.Value = enc_motPasse;
			paramNiveau.Value = niveauAcces;
			paramAbo.Value = listeAbo;

			command1.ExecuteNonQuery();
			connection1.Close();
		}

		private string getEncryptedPass()
		{
			string enc_str = String.Empty;
			SHA512Managed sham = new SHA512Managed();
			Convert.ToBase64String(sham.ComputeHash(Encoding.ASCII.GetBytes(password)));
			byte[] enc_data = ASCIIEncoding.ASCII.GetBytes(password);
			enc_str = Convert.ToBase64String(enc_data);
			return enc_str;
		}

		private string getDecryptedPass()
		{
			string dec_str = String.Empty;
			byte[] dec_data = Convert.FromBase64String(password);
			dec_str = ASCIIEncoding.ASCII.GetString(dec_data);
			return dec_str;
		}

		private string getEncryptedData(string data)
		{
			string enc_str = String.Empty;
			SHA512Managed sham = new SHA512Managed();
			Convert.ToBase64String(sham.ComputeHash(Encoding.ASCII.GetBytes(data)));
			byte[] enc_data = ASCIIEncoding.ASCII.GetBytes(data);
			enc_str = Convert.ToBase64String(enc_data);
			return enc_str;
		}

		private string getDecryptedData(string data)
		{
			string dec_str = String.Empty;
			byte[] dec_data = Convert.FromBase64String(data);
			dec_str = ASCIIEncoding.ASCII.GetString(dec_data);
			return dec_str;
		}

		public string chaineAbo()
		{
			string res = String.Empty;
			for(int i = 0; i < listeAbonnement.Length; i++){
				if(i < listeAbonnement.Length - 1){res = res + listeAbonnement[i] + " ";}else{res = res + listeAbonnement[i];}
			}
			return res;
		}

		public string[] listeTypesEquipement(string domaine, string categorie)
		{
			string[] tab;
			string str = "SELECT TYPE_EQUIPEMENT,DOMAINE,CATEGORIE FROM TABLE_LISTE_TYPE_EQUIPEMENTS WHERE (DOMAINE='" + domaine + "') AND (CATEGORIE='" + categorie + "')";
			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			DataSet ds = new DataSet();
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter (str, connection1);
			oleDbAdapter1.Fill(ds);
			connection1.Close();
			int nombreItems = ds.Tables[0].Rows.Count;
			tab = new string[nombreItems];
			for(int i = 0; i < nombreItems; i++){
				tab[i] = (string) ds.Tables[0].Rows[i]["TYPE_EQUIPEMENT"];
			}
			return tab;
		}

		public string[] listeTypesEquipement(string domaine, string categorie, string espace)
		{
			string[] tab;
			string str = "SELECT TYPE_EQUIPEMENT,DOMAINE,CATEGORIE FROM TABLE_LISTE_TYPE_EQUIPEMENTS WHERE (DOMAINE='" + domaine + "') AND (CATEGORIE='" + categorie + "') AND (ESPACE='" + espace + "')";
			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			DataSet ds = new DataSet();
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter (str, connection1);
			oleDbAdapter1.Fill(ds);
			connection1.Close();
			int nombreItems = ds.Tables[0].Rows.Count;
			tab = new string[nombreItems];
			for(int i = 0; i < nombreItems; i++){
				tab[i] = (string) ds.Tables[0].Rows[i]["TYPE_EQUIPEMENT"];
			}
			return tab;
		}

		public string[] listeTypesEquipement(string categorie)
		{
			string[] tab;
			string str = "SELECT TYPE_EQUIPEMENT,DOMAINE,CATEGORIE FROM TABLE_LISTE_TYPE_EQUIPEMENTS WHERE (CATEGORIE='" + categorie + "') ORDER BY TYPE_EQUIPEMENT";
			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			DataSet ds = new DataSet();
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter (str, connection1);
			oleDbAdapter1.Fill(ds);
			connection1.Close();
			int nombreItems = ds.Tables[0].Rows.Count;
			tab = new string[nombreItems];
			for(int i = 0; i < nombreItems; i++){
				tab[i] = (string) ds.Tables[0].Rows[i]["TYPE_EQUIPEMENT"];
			}
			return tab;
		}

		public string[] listeCategories(string domaine)
		{
			string[] tab;
			string str = "SELECT DOMAINE,CATEGORIE FROM TABLE_LISTE_DOMAINES_CATEGORIES WHERE (DOMAINE='" + domaine + "')";
			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			DataSet ds = new DataSet();
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter (str, connection1);
			oleDbAdapter1.Fill(ds);
			connection1.Close();
			int nombreItems = ds.Tables[0].Rows.Count;
			tab = new string[nombreItems];
			for(int i = 0; i < nombreItems; i++){
				tab[i] = (string) ds.Tables[0].Rows[i]["CATEGORIE"];
			}
			return tab;
		}

		public string[] listeCategories(string domaine, string espace)
		{
			string[] tab;
			string str = "SELECT DOMAINE,CATEGORIE FROM TABLE_LISTE_DOMAINES_CATEGORIES WHERE (DOMAINE='" + domaine + "' AND ESPACE='" + espace + "')";
			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			DataSet ds = new DataSet();
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter (str, connection1);
			oleDbAdapter1.Fill(ds);
			connection1.Close();
			int nombreItems = ds.Tables[0].Rows.Count;
			tab = new string[nombreItems];
			for(int i = 0; i < nombreItems; i++){
				tab[i] = (string) ds.Tables[0].Rows[i]["CATEGORIE"];
			}
			return tab;
		}

		public string[] listeDomaines(string espace)
		{
			string[] tab;
			string str = "SELECT DISTINCT DOMAINE FROM TABLE_LISTE_DOMAINES_CATEGORIES WHERE (ESPACE='" + espace + "')";
			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			DataSet ds = new DataSet();
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter (str, connection1);
			oleDbAdapter1.Fill(ds);
			connection1.Close();
			int nombreItems = ds.Tables[0].Rows.Count;
			tab = new string[nombreItems];
			for(int i = 0; i < nombreItems; i++){
				tab[i] = (string) ds.Tables[0].Rows[i]["DOMAINE"];
			}
			return tab;
		}

		public string[] listeEspaces()
		{
			string[] espaces = new string[]{"MGE", "MMT", "MBE", "MMG", "MEL", "SPE", "", "", "", ""};
			string[] tabAbo = lireAbonnement();
			int taille = 0;
			foreach(string item in tabAbo){
				if(item == "1"){taille = taille + 1;}
			}
			string[] tab = new string[taille];
			int i = 0;
			int index = 0;
			foreach(string item in tabAbo){
				if(item == "1"){tab[i] = espaces[index];i++;}
				index++;
			}
			return tab;
		}

		public string[] tabEspaces()
		{
			string[] espaces = new string[]{"MGE", "MMT", "MBE", "MMG", "MEL", "SPE", "", "", "", ""};
			string[] tabAbo = lireAbonnement();
			string[] tab = new string[10];
			for(int i = 0; i < tab.Length; i++){
				if(tabAbo[i] == "1"){tab[i] = espaces[i];}else{tab[i] = String.Empty;}
			}
			return tab;
		}

		public string creerChaineSQLTypeEquipements()
		{
			string res = String.Empty;
			string[] tabRef = listeEspaces();
			string[] tab = afficheAbonnement;
			string chaineEspace = String.Empty;
			for(int i = 0; i < tabRef.Length; i++){
				if(tab[i] == "1"){
					if(chaineEspace == ""){
						chaineEspace = tabRef[i] + "')";
					}
					else{
						chaineEspace = chaineEspace + " OR (ESPACE='" + tabRef[i] + "')";
					}
				}
			}
			res = "SELECT ESPACE,DOMAINE,CATEGORIE,TYPE,SOUSTYPE_1,SOUSTYPE_2,SOUSTYPE_3,TYPE_EQUIPEMENT FROM TABLE_LISTE_TYPE_EQUIPEMENTS WHERE (ESPACE = '" + chaineEspace + " ORDER BY TYPE_EQUIPEMENT";
			return res;
		}

		public DataSet dsListeTypeEquipements()
		{
			DataSet ds = new DataSet();
			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			string chaineSQL = creerChaineSQLTypeEquipements();
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter (chaineSQL, connection1);
			oleDbAdapter1.Fill(ds);
			connection1.Close();
			return ds;
		}

		public DataSet dsListeTypeEquipements(string chaine)
		{
			DataSet ds = new DataSet();
			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter (chaine, connection1);
			oleDbAdapter1.Fill(ds);
			connection1.Close();
			return ds;
		}

		public DataSet dsListeOperations(string chaine)
		{
			DataSet ds = new DataSet();
			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter (chaine, connection1);
			oleDbAdapter1.Fill(ds);
			connection1.Close();
			return ds;
		}

		public DataSet dsListeCaracteristiques(string chaine)
		{
			DataSet ds = new DataSet();
			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter (chaine, connection1);
			oleDbAdapter1.Fill(ds);
			connection1.Close();
			return ds;
		}

	}

	//-----------------------------------------------------------------------------------------------------

	public class liste_utilisateurs {


	private DataSet ds1;
	private OleDbDataAdapter oleDbAdapter1;
	private OleDbCommandBuilder cb1;
	private string chaineConnection = String.Empty;
	private string nomBaseOperations = String.Empty;


	public liste_utilisateurs(DataSet ds) {

		this.ds1 = ds;

		this.nomBaseOperations = "BASE MAINTENANCE_OPERATIONS ET FMD.mdb";
		this.chaineConnection = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + "\\" + nomBaseOperations + ";Jet OLEDB:Database Password=mexico";;
	}

	private DataSet creerListeUtilisateurs(){

		OleDbConnection connection1 = new OleDbConnection(chaineConnection);
		connection1.Open();
		string strSQL = "SELECT  N°,NOM_UTILISATEUR,MOT_PASSE,NIVEAU_ACCES,LISTE_ABONNEMENT FROM TABLE_UTILISATEURS";
		oleDbAdapter1 = new OleDbDataAdapter (strSQL, connection1);

		oleDbAdapter1.SelectCommand = new OleDbCommand(strSQL, connection1);
        cb1 = new OleDbCommandBuilder(oleDbAdapter1);

		DataSet ds = new DataSet();
		oleDbAdapter1.Fill(ds);
		connection1.Close();
		ds.Tables[0].TableName = "Table UTILISATEURS";
		return ds;
	}

	public void chargerListe() {

		this.ds1 = creerListeUtilisateurs();
	}

	public void majListe() {

		oleDbAdapter1.Update(ds1, "Table UTILISATEURS");
	}

	public DataTable afficheTable(){
		DataTable dt = new DataTable();
		dt = ds1.Tables[0];
		return dt;
	}

	public DataSet afficheDs{
		get {return ds1;}
	}

	public DataSet modifieDs{
		set {this.ds1 = value;}
	}

}

	//-----------------------------------------------------------------------------------------------------

	public class liste_equipements
	{

		private DataTable dt1;
		private DataTable dt2;
		private string chaineConnection = String.Empty;
		private string nomBaseOperations = String.Empty;
		private string[] tableauEcartsTypes;
		private string[] tableauEcartsParam;
		private string[] tableauEcartsClasseG;
		private string[] tableauEcartsReperes;

		private bool test_Equipements = false;
		private bool test_Systemes = false;


		public liste_equipements(DataTable dt) {

			this.dt1 = dt;
			this.dt2 = copieSansValeursNulles(dt1, "REPERE_SECTEUR", "(non défini)");
			this.nomBaseOperations = "BASE MAINTENANCE_OPERATIONS ET FMD.mdb";
			this.chaineConnection = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + "\\" + nomBaseOperations + ";Jet OLEDB:Database Password=mexico";
		}

		public bool testOKEquipements{
			get {return test_Equipements;}
		}

		public bool testOKSystemes{
			get {return test_Systemes;}
		}

		public void chargerListeExcel(string fileExcelName, string type){

			DataSet ds = new DataSet();
			string chaineConnectionExcel = String.Empty;
			string strSQL = "SELECT * FROM [LISTE EQUIPEMENTS$]";
			switch(type){
				case "xls" :
					chaineConnectionExcel = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileExcelName +";Extended Properties=" + (char)34 + "Excel 8.0;HDR=Yes;" + (char)34 + ";";
				break;
				case "xlsx" :
					chaineConnectionExcel = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileExcelName +";Extended Properties=" + (char)34 + "Excel 12.0 Xml;HDR=YES" + (char)34 + ";";
				break;
				default :
					chaineConnectionExcel = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileExcelName +";Extended Properties=" + (char)34 + "Excel 8.0;HDR=Yes;" + (char)34 + ";";
				break;
			}

			OleDbConnection connection1 = new OleDbConnection(chaineConnectionExcel);
			connection1.Open();
			OleDbCommand oleDbCommand1 = new OleDbCommand(strSQL, connection1);
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter();
			oleDbAdapter1.SelectCommand = oleDbCommand1;

			DataTableMapping mapping = oleDbAdapter1.TableMappings.Add("Table", "Table Liste EQUIPEMENTS");
			oleDbAdapter1.Fill(ds);
			connection1.Close();

			// compléter les champs vides
			for(int i = 0; i < ds.Tables[0].Rows.Count; i++){
				if(ds.Tables[0].Rows[i]["TYPE_EQUIPEMENT"].ToString() == ""){ds.Tables[0].Rows[i]["TYPE_EQUIPEMENT"] = "(non identifié)";}
				if(ds.Tables[0].Rows[i]["NOM_SYSTEME"].ToString() == ""){ds.Tables[0].Rows[i]["NOM_SYSTEME"] = "(non affecté)";}
				if(ds.Tables[0].Rows[i]["PARAMETRE_A"].ToString() == ""){ds.Tables[0].Rows[i]["PARAMETRE_A"] = 2;}
				if(ds.Tables[0].Rows[i]["PARAMETRE_B"].ToString() == ""){ds.Tables[0].Rows[i]["PARAMETRE_B"] = 2;}
				if(ds.Tables[0].Rows[i]["PARAMETRE_C"].ToString() == ""){ds.Tables[0].Rows[i]["PARAMETRE_C"] = 2;}
				if(ds.Tables[0].Rows[i]["PARAMETRE_D"].ToString() == ""){ds.Tables[0].Rows[i]["PARAMETRE_D"] = 2;}
				if(ds.Tables[0].Rows[i]["CLASSE_G"].ToString() == ""){ds.Tables[0].Rows[i]["CLASSE_G"] = 0;}
				if(ds.Tables[0].Rows[i]["QUANTITE"].ToString() == ""){ds.Tables[0].Rows[i]["QUANTITE"] = 0;}
				if(ds.Tables[0].Rows[i]["QUANTITE"].ToString() == "0" | ds.Tables[0].Rows[i]["QUANTITE"].ToString() == "1"){
					ds.Tables[0].Rows[i]["CARACTERE"] = "IND";
				}else{
					ds.Tables[0].Rows[i]["CARACTERE"] = "REG";
				}
				if(ds.Tables[0].Rows[i]["LOCALISATION"].ToString() == ""){
					string secteur = ds.Tables[0].Rows[i]["REPERE_SECTEUR"].ToString();
					string batiment = ds.Tables[0].Rows[i]["REPERE_BATIMENT"].ToString();
					string emplacement = ds.Tables[0].Rows[i]["REPERE_EMPLACEMENT"].ToString();
					ds.Tables[0].Rows[i]["LOCALISATION"] = champLocalisation(secteur, batiment, emplacement);
				}
			}
			dt1 = ds.Tables[0];
			dt1.TableName = "Table Liste EQUIPEMENTS";
			oleDbAdapter1.Dispose();
			oleDbCommand1.Dispose();
		}

		public void chargerListeXml(string nomCompletFichierXML) {

			DataSet ds = new DataSet();
			ds.ReadXml(nomCompletFichierXML);
			if(ds.Tables.Contains("Table Liste EQUIPEMENTS") == true){
				dt1 = ds.Tables["Table Liste EQUIPEMENTS"];
			}
		}

		public string champLocalisation(string secteur, string batiment, string emplacement){

			string localisation = String.Empty;
			if(secteur =="" & batiment == "" & emplacement ==""){localisation = "";}
			else{
				if(batiment =="" & emplacement == ""){
						localisation = secteur;
				}else{
					if(emplacement == ""){
						localisation = secteur + " / " + batiment;
					}else{
						localisation = secteur + " / " + batiment + " / " + emplacement;
					}
				}
			}
			return localisation;
		}

		private DataTable dtTypeEquipements(){

			DataTable dt = new DataTable();
			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			DataSet dsTypeEquipements = new DataSet();
			string strTypeEqpts = "SELECT TYPE_EQUIPEMENT FROM TABLE_LISTE_TYPE_EQUIPEMENTS";

			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter (strTypeEqpts, connection1);
			oleDbAdapter1.Fill(dsTypeEquipements);
			dt = dsTypeEquipements.Tables[0];
			connection1.Close();
			return dt;
		}

		private string[] typeRef(DataTable dtTypes){

			int count1 = dtTypes.Rows.Count;
			string[] typesRef = new string[count1];

			for(int i = 0; i < count1; i++){
				DataRow dr = dtTypes.Rows[i];
				typesRef[i] = (string)dr["TYPE_EQUIPEMENT"];
			}
			return typesRef;
		}

		private bool verifierListe(){

			bool ind = false;
			string[] typesRef;
			string[] typesListe;
			int[] listeEcarts;
			int[] listeEcartsClasseG;

			DataTable dtTypes = dtTypeEquipements();
			typesRef = typeRef(dtTypes);

			// vérification de la cohérence des "TYPE_EQUIPEMENT"
			int count2 = dt1.Rows.Count;
			typesListe = new string[count2];
			listeEcarts = new int[count2];
			for(int i = 0; i < count2; i++){
				DataRow dr = dt1.Rows[i];
				try{
					typesListe[i] = (string)dr["TYPE_EQUIPEMENT"];
				}catch{
					typesListe[i] = "";
				}
			}

			bool indEcartTypesEqpts = testEcart(typesListe, typesRef);
			listeEcarts = creerListeEcarts(typesListe, typesRef);
			tableauEcartsTypes = new string[listeEcarts.Length];

			if(indEcartTypesEqpts == true){

				int k = 0;
				do{
					tableauEcartsTypes[k] = "Ligne " + dt1.Rows[listeEcarts[k]]["NUM"].ToString() + "    " + (string)dt1.Rows[listeEcarts[k]]["REPERE_EQUIPEMENT"];
					k++;
				}
				while(listeEcarts[k] !=0);
			}

			// vérification des paramètres
			double[] paramAListe = new double[count2];
			double[] paramBListe = new double[count2];
			double[] paramCListe = new double[count2];
			double[] paramDListe = new double[count2];
			double[] listeEcartsParam = new double[count2];

			for(int p = 0; p < count2; p++){
				DataRow dr = dt1.Rows[p];
				paramAListe[p] = double.Parse(dr["PARAMETRE_A"].ToString());
				paramBListe[p] = double.Parse(dr["PARAMETRE_B"].ToString());
				paramCListe[p] = double.Parse(dr["PARAMETRE_C"].ToString());
				paramDListe[p] = double.Parse(dr["PARAMETRE_D"].ToString());
			}

			bool indEcartsParam = testEcartParam(paramAListe) | testEcartParam(paramBListe) | testEcartParam(paramCListe) | testEcartParam(paramDListe);

			Array array1, array2, array3, array4, arrayComplet;
			array1 = creerListeEcartsParam(paramAListe);
			array2 = creerListeEcartsParam(paramBListe);
			array3 = creerListeEcartsParam(paramCListe);
			array4 = creerListeEcartsParam(paramDListe);
			arrayComplet = Array.CreateInstance(typeof(Int32), array1.Length + array2.Length + array3.Length + array4.Length);
			array1.CopyTo(arrayComplet, 0);
			array2.CopyTo(arrayComplet, array1.Length);
			array3.CopyTo(arrayComplet, array1.Length + array2.Length);
			array4.CopyTo(arrayComplet, array1.Length + array2.Length + array3.Length);
			Array.Sort(arrayComplet);
			Array.Reverse(arrayComplet);

			int[] listeIndexEcartsParam = new int[arrayComplet.Length];

			for(int i = 0; i < listeIndexEcartsParam.Length; i++){
				listeIndexEcartsParam[i] = (int) arrayComplet.GetValue(i);
			}

			tableauEcartsParam = new string[listeIndexEcartsParam.Length];

			if(indEcartsParam == true){
				for(int m = 0; m < listeIndexEcartsParam.Length; m++){
					if(listeIndexEcartsParam[m] != -1){
						tableauEcartsParam[m] = "Ligne " + dt1.Rows[listeIndexEcartsParam[m]]["NUM"].ToString() + "    " + (string)dt1.Rows[listeIndexEcartsParam[m]]["REPERE_EQUIPEMENT"];
					}
				}
			}

			// vérification des classe_G
			double[] classeGListe = new double[count2];
			for(int p = 0; p < count2; p++){

				DataRow dr = dt1.Rows[p];
				classeGListe[p] = double.Parse(dr["CLASSE_G"].ToString());
			}

			bool indEcartsClasseG = testEcartClasseG(classeGListe);
			listeEcartsClasseG = creerListeEcartsClasseG(classeGListe);

			Array.Sort(listeEcartsClasseG);
			Array.Reverse(listeEcartsClasseG);

			tableauEcartsClasseG = new string[listeEcartsClasseG.Length];

			if(indEcartsClasseG == true){
				for(int m = 0; m < listeEcartsClasseG.Length; m++){
					if(listeEcartsClasseG[m] != -1){
						tableauEcartsClasseG[m] =  "Ligne " + dt1.Rows[listeEcartsClasseG[m]]["NUM"].ToString() + "    " + (string)dt1.Rows[listeEcartsClasseG[m]]["REPERE_EQUIPEMENT"];
					}
				}
			}

			// vérification des doublons REPERE_EQUIPEMENT
			string[] listeReperes = new string[count2];
			for(int i = 0; i < count2; i++){
				DataRow dr = dt1.Rows[i];
				try{
					listeReperes[i] = (string)dr["REPERE_EQUIPEMENT"];
				}catch{
					listeReperes[i] = "";
				}
			}
			bool indEcartsReperes = testEcartReperes(listeReperes);
			int[] listeEcartsReperes = creerListeEcartsReperes(listeReperes);
			tableauEcartsReperes = new string[listeEcartsReperes.Length];

			if(indEcartsReperes == true){
				int k = 0;
				do{
					tableauEcartsReperes[k] = "Ligne " + dt1.Rows[listeEcartsReperes[k]]["NUM"].ToString() + "    " + (string)dt1.Rows[listeEcartsReperes[k]]["REPERE_EQUIPEMENT"];
					k++;
				}
				while(listeEcartsReperes[k] != 0);
			}

			// posiitonner l'indicateur de conformité global
			ind = !indEcartTypesEqpts && !indEcartsParam && !indEcartsClasseG && !indEcartsReperes;
			return ind;
		}

		private bool verifierListeSystemes()
		{
			bool res = false; // true pour OK
			string[] tabSystemes = listeSystemes();
			if(tabSystemes.Length <= 10){
				res = true;
			}
			return res;
		}

		public bool verifierGlobal()
		{
			bool res = false;
			bool ind1 = verifierListe();
			bool ind2 = verifierListeSystemes();
			if(ind1 == true){this.test_Equipements = true;}else{this.test_Equipements = false;}
			if(ind2 == true){this.test_Systemes = true;}else{this.test_Systemes = false;}
			res = ind1 && ind2;
			return res;
		}

		private bool testEcartParam(double[] listeParam)
		{
			bool indTest = false;
			int rang = 0;
			double[] tabParam = new double[]{0, 1, 2, 3, 4};
			foreach(double item in listeParam){
				int test = Array.IndexOf(tabParam, item);
				if(test == -1){
					indTest = true;
				}
				rang++;
			}
			return indTest;
		}

		private bool testEcartClasseG(double[] listeClasseG)
		{
			bool indTest = false;
			int rang = 0;
			double[] tabClasseG = new double[]{0, 1, 2};
			foreach(double item in listeClasseG){
				int test = Array.IndexOf(tabClasseG, item);
				if(test == -1){
					indTest = true;
				}
				rang++;
			}
			return indTest;
		}

		private bool testEcartReperes(string[] listeRep)
		{
			bool indTest = false;
			foreach(string item in listeRep){
				int rang = Array.IndexOf(listeRep, item);
				for(int i = 0; i < rang; i++){
					if(item.ToUpper() == listeRep[i].ToUpper()){indTest = true;}
				}
				for(int j = rang + 1; j < listeRep.Length; j++){
					if(item.ToUpper() == listeRep[j].ToUpper()){indTest = true;}
				}
			}
			return indTest;
		}

		public void verifierReperes()
		{
			// compléter les champs REPERE_EQUIPEMENT vides
			int numero = 0;
			foreach(DataRow dr in dt1.Rows){
				if(dt1.Rows[numero]["REPERE_EQUIPEMENT"] == DBNull.Value || (string)dt1.Rows[numero]["REPERE_EQUIPEMENT"] == String.Empty){
					dt1.Rows[numero]["REPERE_EQUIPEMENT"] = "REP_" + dt1.Rows[numero]["NUM"].ToString();
				}
			numero++;
			}
		}

		private int[] creerListeEcartsParam(double[] listeTest)
		{
			int[] listeEcarts = new int[listeTest.Length];
			double[] tabParam = new double[]{0, 1, 2, 3, 4};
			int rang = 0;
			int i = 0;
			foreach(double item in listeTest){
				int test = Array.IndexOf(tabParam, item);
				if(test == -1){
					listeEcarts[i] = rang;
				}else{
					listeEcarts[i] = -1;
				}
				i++;
				rang++;
			}
			return listeEcarts;
		}

		private int[] creerListeEcartsReperes(string[] listeRep)
		{
			int[] listeEcarts = new int[listeRep.Length];
			int index = 0;
			foreach(string item in listeRep){
				int rang = Array.IndexOf(listeRep, item);
				for(int i = 0; i < rang; i++){
					if(item.ToUpper() == listeRep[i].ToUpper()){listeEcarts[index] = i;index++;}
				}
				for(int j = rang + 1; j < listeRep.Length; j++){
					if(item.ToUpper() == listeRep[j].ToUpper()){listeEcarts[index] = j;index++;}
				}
			}
			return listeEcarts;
		}

		private int[] creerListeEcartsClasseG(double[] listeTest)
		{
			int[] listeEcarts = new int[listeTest.Length];
			double[] tabClasseG = new double[]{0, 1, 2};
			int rang = 0;
			int i = 0;
			foreach(double item in listeTest){
				int test = Array.IndexOf(tabClasseG, item);
				if(test == -1){
					listeEcarts[i] = rang;
				}else{
					listeEcarts[i] = -1;
				}
				i++;
				rang++;
			}
			return listeEcarts;
		}

		private int[] creerListeEcarts(string[] listeTest, string[] listeRef)
		{
			int[] listeEcarts = new int[listeRef.Length];
			int rang = 0;
			int i = 0;
			foreach(string item in listeTest){
				if(item != ""){
					int test = Array.IndexOf(listeRef, item);
					if(test == -1){
						listeEcarts[i] = rang;
						i++;
					}
					rang++;
				}else{
					listeEcarts[i] = rang;
					i++;
					rang++;
				}
			}
			return listeEcarts;
		}

		private bool testEcart(string[] listeTest, string[] listeRef)
		{
			bool indTest = false;
			int rang = 0;
			foreach(string item in listeTest){
				if(item != ""){
					int test = Array.IndexOf(listeRef, item);
					if(test == -1){
						indTest = true;
					}
					rang++;
				}else{
					indTest = true;
				}
			}
			return indTest;
		}

		private string listeErreurs(string[] tab)
		{
			int i = 0;
			string text = String.Empty;
			while(tab[i] != null){
				text = text + tab[i] + "\n";
				i++;
			}
			return text;
		}

		public string chaineErreurs()
		{
			string chaine = String.Empty;
			string text1 = String.Empty;
			string text2 = String.Empty;
			string text3 = String.Empty;
			string text4 = String.Empty;
			if(listeErreurs(tabEcartsTypes) != String.Empty){
				text1 = "LIGNES CONTENANT DES ERREURS SUR LES TYPE_EQUIPEMENT :\n" + listeErreurs(tabEcartsTypes);
			}
			if(listeErreurs(tabEcartsReperes) != String.Empty){
				text2 = "LIGNES CONTENANT DES ERREURS SUR LES REPERES :\n" + listeErreurs(tabEcartsReperes);
			}
			if(listeErreurs(tabEcartsParam) != String.Empty){
				text3 = "LIGNES CONTENANT DES ERREURS SUR LES PARAMETRES :\n" + listeErreurs(tabEcartsParam);
			}
			if(listeErreurs(tabEcartsClasseG) != String.Empty){
				text4 = "LIGNES CONTENANT DES ERREURS SUR LES CLASSE_G :\n" + listeErreurs(tabEcartsClasseG);
			}
			return chaine = (text1 + "\n" + text2 + "\n" +  text3 + "\n" + text4);
		}

		public double nombreTotalEquipements(bool perim)
		{
			double res = 0;
			string selection = "APPARTENANCE_PERIMETRE = 'OUI'" + "OR APPARTENANCE_PERIMETRE = 'NON'";
			if(perim == true){
				selection = "APPARTENANCE_PERIMETRE ='" + "OUI" + "'";
			}
			DataRow[] dr1 = dt1.Select(selection);
			for(int i = 0; i < dr1.Length; i++){
				 res =  res + double.Parse(dr1[i]["QUANTITE"].ToString());
			}
			return res;
		}

		public double nombreTotalEquipements(int classe)
		{
			double res = 0;
			string selection = "CLASSE_G ='" + classe.ToString() + "'";
			DataRow[] dr1 = dt1.Select(selection);
			for(int i = 0; i < dr1.Length; i++){
				 res =  res + double.Parse(dr1[i]["QUANTITE"].ToString());
			}
			return res;
		}

		public double nombreTotalEquipements(int classe, bool perim)
		{
			double res = 0;
			string chainePerimetre = String.Empty;
			if(perim == true){
				chainePerimetre = "AND APPARTENANCE_PERIMETRE ='" + "OUI" + "'";
			}
			string selection = "CLASSE_G ='" + classe.ToString() + "'" + chainePerimetre;
			DataRow[] dr1 = dt1.Select(selection);
			for(int i = 0; i < dr1.Length; i++){
				 res =  res + double.Parse(dr1[i]["QUANTITE"].ToString());
			}
			return res;
		}

		public double nombreEquipementsParSysteme(int classe, string syst)
		{
			double res = 0;
			string chaineSysteme = "AND NOM_SYSTEME ='" + syst + "'";
			string selection = "CLASSE_G ='" + classe.ToString() + "'" + chaineSysteme;
			DataRow[] dr1 = dt1.Select(selection);
			for(int i = 0; i < dr1.Length; i++){
				 res =  res + double.Parse(dr1[i]["QUANTITE"].ToString());
			}
			return res;
		}

		public double nombreEquipementsParSysteme(int classe, string syst, bool perim)
		{
			double res = 0;
			string chaineSysteme = "AND NOM_SYSTEME ='" + syst;
			string chainePerimetre = String.Empty;
			if(perim == true){
				chainePerimetre = "AND APPARTENANCE_PERIMETRE ='" + "OUI" + "'";
			}
			string selection = "CLASSE_G ='" + classe.ToString() + "'" + chaineSysteme + "'" + chainePerimetre;
			DataRow[] dr1 = dt1.Select(selection);
			for(int i = 0; i < dr1.Length; i++){
				 res =  res + double.Parse(dr1[i]["QUANTITE"].ToString());
			}
			return res;
		}

		public double nombreEquipementsParSysteme(string systeme)
		{
			double res = 0;
			string selection = "NOM_SYSTEME ='" + systeme.ToString() + "'";
			DataRow[] dr1 = dt1.Select(selection);
			for(int i = 0; i < dr1.Length; i++){
				 res =  res + double.Parse(dr1[i]["QUANTITE"].ToString());
			}
			return res;
		}

		public double nombreEquipementsParSysteme(string systeme, bool perim)
		{
			double res = 0;
			string chainePerimetre = String.Empty;
			if(perim == true){
				chainePerimetre = "AND APPARTENANCE_PERIMETRE ='" + "OUI" + "'";
			}
			string selection = "NOM_SYSTEME ='" + systeme.ToString() + "'" + chainePerimetre;
			DataRow[] dr1 = dt1.Select(selection);
			for(int i = 0; i < dr1.Length; i++){
				 res =  res + double.Parse(dr1[i]["QUANTITE"].ToString());
			}
			return res;
		}

		public double nombreEquipementsParSecteur(int classe, string secteur)
		{
			double res = 0;
			string chaineSecteur = "AND REPERE_SECTEUR ='" + secteur + "'";
			string selection = "CLASSE_G ='" + classe.ToString() + "'" + chaineSecteur;
			DataRow[] dr1 = dt2.Select(selection);
			for(int i = 0; i < dr1.Length; i++){
				 res =  res + double.Parse(dr1[i]["QUANTITE"].ToString());
			}
			return res;
		}

		public double nombreEquipementsParSecteur(int classe, string secteur, bool perim)
		{
			double res = 0;
			string chaineSecteur = "AND REPERE_SECTEUR ='" + secteur;
			string chainePerimetre = String.Empty;
			if(perim == true){
				chainePerimetre = "AND APPARTENANCE_PERIMETRE ='" + "OUI" + "'";
			}
			string selection = "CLASSE_G ='" + classe.ToString() + "'" + chaineSecteur + "'" + chainePerimetre;
			DataRow[] dr1 = dt2.Select(selection);
			for(int i = 0; i < dr1.Length; i++){
				 res =  res + double.Parse(dr1[i]["QUANTITE"].ToString());
			}
			return res;
		}

		private DataTable copieSansValeursNulles(DataTable dt, string champ, string valeur)
		{
			DataTable dt0 = dt.Copy();
			for(int i = 0; i < dt1.Rows.Count; i++){
				if(dt0.Rows[i][champ].ToString() == String.Empty){dt0.Rows[i][champ] = valeur;}
			}
			return dt0;
		}

		public double nombreEquipementsParSecteur(string secteur)
		{
			double res = 0;
			string selection = "REPERE_SECTEUR ='" + secteur.ToString() + "'";
			DataRow[] dr1 = dt2.Select(selection);
			for(int i = 0; i < dr1.Length; i++){
				 res =  res + double.Parse(dr1[i]["QUANTITE"].ToString());
			}
			return res;
		}

		public double nombreEquipementsParSecteur(string secteur, bool perim)
		{
			double res = 0;
			string chainePerimetre = String.Empty;
			if(perim == true){
				chainePerimetre = "AND APPARTENANCE_PERIMETRE ='" + "OUI" + "'";
			}
			string selection = "REPERE_SECTEUR ='" + secteur.ToString() + "'" + chainePerimetre;
			DataRow[] dr1 = dt2.Select(selection);
			for(int i = 0; i < dr1.Length; i++){
				 res =  res + double.Parse(dr1[i]["QUANTITE"].ToString());
			}
			return res;
		}

		public double[] quantitesParClasse()
		{
			double[] tab = new double[3];
			tab[0] = nombreTotalEquipements(0);
			tab[1] = nombreTotalEquipements(1);
			tab[2] = nombreTotalEquipements(2);
			return tab;
		}

		public double[] quantitesParClasse(bool perim)
		{
			double[] tab = new double[3];
			tab[0] = nombreTotalEquipements(0, perim);
			tab[1] = nombreTotalEquipements(1, perim);
			tab[2] = nombreTotalEquipements(2, perim);
			return tab;
		}

		public string[] listeReperes()
		{
			DataTable dt = selectDistinct("", dt1, "REPERE_EQUIPEMENT");
			string[] tab = new string[dt.Rows.Count];
			int i = 0;
			foreach(DataRow r in dt.Rows){
				tab[i] = r["REPERE_EQUIPEMENT"].ToString();
				i++;
			}
			return tab;
		}

		public string[] listeSystemes()
		{
			DataTable dt = selectDistinct("", dt1, "NOM_SYSTEME");
			string[] tab = new string[dt.Rows.Count];
			int i = 0;
			foreach(DataRow r in dt.Rows){
				tab[i] = r["NOM_SYSTEME"].ToString();
				i++;
			}
			return tab;
		}

		public string[] listeTypeEquipements()
		{
			DataTable dt = selectDistinct("", dt1, "TYPE_EQUIPEMENT");
			string[] tab = new string[dt.Rows.Count];
			int i = 0;
			foreach(DataRow r in dt.Rows){
				tab[i] = r["TYPE_EQUIPEMENT"].ToString();
				i++;
			}
			return tab;
		}

		private string creerChaineTypesEquipements()
		{
			string chaine = "(";
			for(int i = 0; i < listeTypeEquipements().Length; i++){
				chaine = chaine + "'" + listeTypeEquipements()[i] + "'";
				if(i < listeTypeEquipements().Length - 1){chaine = chaine + ",";}
			}
			chaine = chaine + ")";
			return chaine;
		}

		public DataTable tableTypologie()
		{
			DataTable dt = new DataTable();
			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			DataSet dsTypeEquipements = new DataSet();
			string chaine = creerChaineTypesEquipements();
			string strTypeEqpts = "SELECT TYPE_EQUIPEMENT,ESPACE,DOMAINE,CATEGORIE FROM Liste_CARACTERISTIQUES_FMD WHERE TYPE_EQUIPEMENT IN " + chaine;

			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter (strTypeEqpts, connection1);
			oleDbAdapter1.Fill(dsTypeEquipements);
			dt = dsTypeEquipements.Tables[0];
			connection1.Close();
			return dt;
		}

		public DataTable tableFMD()
		{
			DataTable dt = new DataTable();
			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			DataSet dsTypeEquipements = new DataSet();
			string chaine = creerChaineTypesEquipements();
			string strTypeEqpts = "SELECT TYPE_EQUIPEMENT,BETA,ETA,GAMMA,TD_NORM_PAR_HEURE,FMAD,N_DEF_PAR_AN,TMRS_MC FROM Liste_CARACTERISTIQUES_FMD WHERE TYPE_EQUIPEMENT IN " + chaine;

			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter (strTypeEqpts, connection1);
			oleDbAdapter1.Fill(dsTypeEquipements);
			dt = dsTypeEquipements.Tables[0];
			connection1.Close();
			return dt;
		}

		private DataTable selectDistinct(string nomTable, DataTable tableSource, string champ)
		{
		   DataTable dt = new DataTable(nomTable);
		   dt.Columns.Add(champ, tableSource.Columns[champ].DataType);
		   object LastValue = null;
		   foreach (DataRow dr in tableSource.Select("", champ)){
			   if(LastValue == null || !(DataColumn.Equals(LastValue, dr[champ]))){
				  LastValue = dr[champ];
				  dt.Rows.Add(new object[]{LastValue});
			   }
		   }
		   return dt;
		}

		public string[] tabEcartsTypes{
			get {return tableauEcartsTypes;}
		}

		public string[] tabEcartsParam{
			get {return tableauEcartsParam;}
		}

		public string[] tabEcartsClasseG{
			get {return tableauEcartsClasseG;}
		}

		public string[] tabEcartsReperes{
			get {return tableauEcartsReperes;}
		}

		public DataTable  afficheDt{
			get {return dt1;}
		}
	}

	//-----------------------------------------------------------------------------------------------------

	public class liste_plan {


		private DataTable dt1;
		private string chaineConnection = String.Empty;
		private string nomBaseOperations = String.Empty;


		public liste_plan(DataTable dt) {

			this.dt1 = dt;

			this.nomBaseOperations = "BASE MAINTENANCE_OPERATIONS ET FMD.mdb";
			this.chaineConnection = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + "\\" + nomBaseOperations + ";Jet OLEDB:Database Password=mexico";;
		}

		private DataTable dtTypeEquipements(){

			DataTable dt = new DataTable();
			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			DataSet dsTypeEquipements = new DataSet();
			string strTypeEqpts = "SELECT TYPE_EQUIPEMENT FROM TABLE_LISTE_TYPE_EQUIPEMENTS";

			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter (strTypeEqpts, connection1);
			oleDbAdapter1.Fill(dsTypeEquipements);
			dt = dsTypeEquipements.Tables[0];
			connection1.Close();
			return dt;
		}

		public string[] listeSystemes(){

			DataTable dt = selectDistinct("", dt1, "NOM_SYSTEME");
			string[] tab = new string[dt.Rows.Count];
			int i = 0;
			foreach(DataRow r in dt.Rows){
				tab[i] = r["NOM_SYSTEME"].ToString();
				i++;
			}
			return tab;
		}

		private double frequence(operation op){

			double  res = 0;
			double frequence = 0;
			double periodicite = double.Parse(op.afficheCar(10));
			switch (op.afficheCar(11)){
				case "AN" :
					frequence = 1 / periodicite;
				break;
				case "SEMESTRE" :
					frequence = 2 / periodicite;
				break;
				case "TRIMESTRE" :
					frequence = 4 / periodicite;
				break;
				case "MOIS" :
					frequence = 12 / periodicite;
				break;
				case "SEMAINE" :
					frequence = 52 / periodicite;
				break;
				case "JOUR" :
					frequence = 365 / periodicite;
				break;
				case "HEURE DE SERVICE" :
					frequence = 8000 / periodicite;
				break;
				default :
					frequence = 12 / periodicite;
				break;
			}
			res =  Math.Round(double.Parse(op.afficheCar(12)) * frequence, 0);
			return res;
		}

		private operation creerOperation(DataRow dr){

			operation op = new operation();
			string typ = String.Empty;
			string[] tabCar = new string[20];
			string code = dr["CODE_OPERATION"].ToString();
			tabCar[0] = dr["TYPE_EQUIPEMENT"].ToString();
			tabCar[10] = dr["PERIODICITE"].ToString();
			tabCar[11] = dr["UNITE_PERIODICITE"].ToString();
			tabCar[12] = dr["QUANTITE"].ToString();
			gamme g = new gamme();
			op.initialiseOperation(code, typ, tabCar, g, nomBaseOperations);
			return op;
		}

		private operation[] listeOperations(){

			operation[] listOp = new operation[dt1.Rows.Count];
			int i = 0;
			foreach(DataRow dr in dt1.Rows){
				operation op = creerOperation(dr);
				listOp[i] = op;
				i++;
			}
			return listOp;
		}

		private operation[] listeOperations(string systeme){

			string selection = "NOM_SYSTEME ='" + systeme.ToString() + "'";
			DataRow[] dr1 = dt1.Select(selection);
			operation[] listOp = new operation[dr1.Length];
			int i = 0;
			foreach(DataRow dr in dr1){
				operation op = creerOperation(dr);
				listOp[i] = op;
				i++;
			}
			return listOp;
		}

		public double nombreTotalOperations(){

			double  res = 0;
			foreach(operation op in listeOperations()){
				res = res + frequence(op);
			}
			return res;
		}

		public double nombreTotalOperations(string systeme){

			double  res = 0;
			foreach(operation op in listeOperations(systeme)){
				res = res + frequence(op);
			}
			return res;
		}

		public string[] listeCodesOperations(){

			DataTable dt = selectDistinct("", dt1, "CODE_OPERATION");
			string[] tab = new string[dt.Rows.Count];
			int i = 0;
			foreach(DataRow r in dt.Rows){
				tab[i] = r["CODE_OPERATION"].ToString();
				i++;
			}
			return tab;
		}

		public string[] listeTypeEquipements(){

			DataTable dt = selectDistinct("", dt1, "TYPE_EQUIPEMENT");
			string[] tab = new string[dt.Rows.Count];
			int i = 0;
			foreach(DataRow r in dt.Rows){
				tab[i] = r["TYPE_EQUIPEMENT"].ToString();
				i++;
			}
			return tab;
		}

		private DataTable selectDistinct(string nomTable, DataTable tableSource, string champ){

		   DataTable dt = new DataTable(nomTable);
		   dt.Columns.Add(champ, tableSource.Columns[champ].DataType);

		   object LastValue = null;
		   foreach (DataRow dr in tableSource.Select("", champ)){
			   if(LastValue == null || !(DataColumn.Equals(LastValue, dr[champ]))){
				  LastValue = dr[champ];
				  dt.Rows.Add(new object[]{LastValue});
			   }
		   }
		   return dt;
		}
	}

	//-----------------------------------------------------------------------------------------------------

	public class equipement {

			private string repere;
			private string[] tableauCaracteristiques;
			private int[] tableauParametres;
			private typEtat _etat;

			private double[,] tabCoeffCor = new double[,] {{0.70, 0.85, 1.00, 1.15, 1.30},{0.85, 0.90, 1.00, 1.10, 1.15},{1.50, 1.25, 1.00, 0.90, 0.80},{1.20, 1.10, 1.00, 0.90, 0.80}};

			public enum typEtat { vide, configure, modifie, enregistre, enSelection };



			public void initialiseEquipement(string rep, string[] TabCar, int[] TabParam) {

				this.repere = rep;
				this.tableauCaracteristiques = TabCar;
				this.tableauParametres = TabParam;
				this._etat = typEtat.vide;
			}

			public string afficheRepere{
				get {return repere;}
			}

			public typEtat etat{
				get {return _etat;}
				set {this._etat = value;}
			}

			public string afficheCar(int n){
				return tableauCaracteristiques[n];
			}

			public int afficheParamEqpt(int n){
				return tableauParametres[n];
			}

			public string[] afficheTabCar{
				get {return tableauCaracteristiques;}
			}

			public int[] afficheTabParam{
				get {return tableauParametres;}
			}

			public void modifieRepere(string _rep){
				this.repere = _rep;
			}

			public void modifieParamEqpt(int n, int value){
				this.tableauParametres[n] = value;
			}

			public void modifieTableauParam(int[] tab){
				this.tableauParametres = tab;
			}

			public void modifieCar(int n, string value){
				this.tableauCaracteristiques[n] = value;
			}

			public void modifieTableauCar(string[] tab){
				this.tableauCaracteristiques = tab;
			}

			public bool testParam(){
				bool res = false;
				res = testParam(0) && testParam(1) && testParam(2) && testParam(3);
				return res;
			}

			public bool testParam(int n){
				bool res = false;
				if(afficheParamEqpt(n) >-1 && afficheParamEqpt(n) <=4){res = true;}
				return res;
			}

			public double calculCoefficientCorrecteur(equipement E){
				double result;
				double kA, kB, kC, kD;

				kA = tabCoeffCor[0, E.afficheParamEqpt(0)];
				kB = tabCoeffCor[1, E.afficheParamEqpt(1)];
				kC = tabCoeffCor[2, E.afficheParamEqpt(2)];
				kD = tabCoeffCor[3, E.afficheParamEqpt(3)];

				result = kA * kB * kC * kD;
				return result;
			}

			public double calculCoefficientKp1(equipement E){
				double result;
				result = tabCoeffCor[1, E.afficheParamEqpt(1)];
				return result;
			}

			public double calculCoefficientKp2(equipement E){
				double result;
				result = tabCoeffCor[3, E.afficheParamEqpt(3)];
				return result;
			}

			public string champLocalisation(){

				string localisation = String.Empty;
				string secteur = afficheCar(4);
				string batiment = afficheCar(5);
				string emplacement = afficheCar(6);

				if(secteur =="" & batiment == "" & emplacement ==""){localisation = "";}
				else{
					if(batiment =="" & emplacement == ""){
							localisation = secteur;
					}
					else{
						if(emplacement == ""){
						localisation = secteur + " / " + batiment;
					}
						else{
							localisation = secteur + " / " + batiment + " / " + emplacement;
						}
					}
				}
				return localisation;
			}

			private string[] champsLocalisation(string loc)
			{
				string[] res = new string[4];
				string[] valeurs = loc.Split(new Char[] {'/'});
				if(valeurs.Length > 2){
						res[0] = valeurs[0];
						res[1] = valeurs[1];
						res[2] = valeurs[2];
				}else{
					if(valeurs.Length > 1){
						res[0] = valeurs[0];
						res[1] = valeurs[1];
					}else{
						res[0] = valeurs[0];
					}
				}
				return res;
			}

			public void configurer(DataRow dr)
			{
				repere = dr["REPERE_EQUIPEMENT"].ToString();

				//initialisation tableau Caractéristiques
				tableauCaracteristiques[0] = dr["TYPE_EQUIPEMENT"].ToString();
				tableauCaracteristiques[1] = dr["DESIGNATION"].ToString();
				tableauCaracteristiques[2] = dr["NOM_SYSTEME"].ToString();
				tableauCaracteristiques[3] = dr["MARQUE_TYPE"].ToString();
				tableauCaracteristiques[4] = champsLocalisation(dr["LOCALISATION"].ToString())[0];
				tableauCaracteristiques[5] = champsLocalisation(dr["LOCALISATION"].ToString())[1];
				tableauCaracteristiques[6] = champsLocalisation(dr["LOCALISATION"].ToString())[2];
				tableauCaracteristiques[7] = champsLocalisation(dr["LOCALISATION"].ToString())[3];
				tableauCaracteristiques[8] = dr["ANNEE_MISE_EN_SERVICE"].ToString();
				tableauCaracteristiques[9] = dr["CLASSE_G"].ToString();
				tableauCaracteristiques[10] = rechercheCaractere(dr["QUANTITE"].ToString());
				tableauCaracteristiques[11] = dr["QUANTITE"].ToString();
				// ...
				tableauCaracteristiques[17] = dr["APPARTENANCE_PERIMETRE"].ToString();
				tableauCaracteristiques[29] = dr["NUM"].ToString();

				//initialisation tableau Paramètres Environnement
				tableauParametres[0] = int.Parse(dr["PARAMETRE_A"].ToString());
				tableauParametres[1] = int.Parse(dr["PARAMETRE_B"].ToString());
				tableauParametres[2] = int.Parse(dr["PARAMETRE_C"].ToString());
				tableauParametres[3] = int.Parse(dr["PARAMETRE_D"].ToString());
			}

			private string rechercheCaractere(string quant)
			{
				string res = String.Empty;
				if(quant != "1"){
					res = "IND";
				}else{
					res = "REG";
				}
				return res;
			}
	}

// Définition de tableauCaracteristiques EQUIPEMENT

// tableauCaracteristiques[0]		TYPE_EQUIPEMENT
// tableauCaracteristiques[1]		DESIGNATION
// tableauCaracteristiques[2]		NOM_SYSTEME
// tableauCaracteristiques[3]		MARQUE_TYPE
// tableauCaracteristiques[4]		REPERE_SECTEUR
// tableauCaracteristiques[5]		REPERE_BATIMENT
// tableauCaracteristiques[6]		REPERE_EMPLACEMENT
// tableauCaracteristiques[7]		LOCALISATION
// tableauCaracteristiques[8]		ANNEE_MISE_EN_SERVICE
// tableauCaracteristiques[9]		CLASSE_G
// tableauCaracteristiques[10]		CARACTERE
// tableauCaracteristiques[11]		QUANTITE
// tableauCaracteristiques[12]		CHAMP_PARAMETRABLE_1
// tableauCaracteristiques[13]		CHAMP_PARAMETRABLE_2
// tableauCaracteristiques[14]		CHAMP_PARAMETRABLE_3
// tableauCaracteristiques[15]		CHAMP_PARAMETRABLE_4
// tableauCaracteristiques[16]		CHAMP_PARAMETRABLE_5
// tableauCaracteristiques[17]		APPARTENANCE_PERIMETRE


// tableauCaracteristiques[29]		numéro ordre


	//-----------------------------------------------------------------------------------------------------

	public class typeEquipement {

		private string nomType;
		private string[] tableauTypologie;
		private string[] tableauCaracteristiquesFMD;
		private string nomBaseOperations;
		private string chaineConnection;

		public void initialiseTypeEquipement(string baseName, string typName, string[] tabTypo, string[] tabCarFMD) {

			this.nomBaseOperations = baseName;
			this.nomType = typName;
			this.tableauTypologie = tabTypo;
			this.tableauCaracteristiquesFMD = tabCarFMD;
			this.chaineConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + "\\" + nomBaseOperations + ";Jet OLEDB:Database Password=mexico";
		}

		public string afficheType{
			get {return nomType;}
		}

		public string afficheTypologie(int n){
			return tableauTypologie[n];
		}

		public string[] afficheTabTypologie{
			get {return tableauTypologie;}
		}

		public string[] afficheTabCaracteristiquesFMD{
			get {return tableauCaracteristiquesFMD;}
		}

		public string afficheCarFMD(int n){
			return tableauCaracteristiquesFMD[n];
		}

		public void modifieType(string _type){
			this.nomType = _type;
		}

		public void modifieTypologie(int n, string value){
			this.tableauTypologie[n] = value;
		}

		public void modifieTableauTypologie(string[] tab){
			this.tableauTypologie = tab;
		}

		public void modifieCarFMD(int n, string value){
			this.tableauCaracteristiquesFMD[n] = value;
		}

		public void modifieTableauCaracteristiquesFMD(string[] tab){
			this.tableauCaracteristiquesFMD = tab;
		}

		public string[] chargerTypologie() {

			string[] tab = new string[10];
			string strSQL = "SELECT ESPACE, DOMAINE, CATEGORIE, TYPE, SOUSTYPE_1, SOUSTYPE_2, SOUSTYPE_3, TYPE_EQUIPEMENT FROM TABLE_LISTE_TYPE_EQUIPEMENTS WHERE (TYPE_EQUIPEMENT= '" + nomType + "')";
	 		OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter(strSQL, connection1);
			connection1.Close();
			DataSet dsTypologie = new DataSet();
			oleDbAdapter1.Fill(dsTypologie);

			DataTable dt1 = dsTypologie.Tables[0];

			if(dt1.Rows.Count != 0){
				tab[0] = dt1.Rows[0]["ESPACE"].ToString();
				tab[1] = dt1.Rows[0]["DOMAINE"].ToString();
				tab[2] = dt1.Rows[0]["CATEGORIE"].ToString();
				tab[3] = dt1.Rows[0]["TYPE"].ToString();
				tab[4] = dt1.Rows[0]["SOUSTYPE_1"].ToString();
				tab[5] = dt1.Rows[0]["SOUSTYPE_2"].ToString();
				tab[6] = dt1.Rows[0]["SOUSTYPE_3"].ToString();
				tab[7] = dt1.Rows[0]["TYPE_EQUIPEMENT"].ToString();
			}
			else{
				for(int i = 0; i < tab.Length; i++){tab[i] = String.Empty;}
			}
			return tab;
		}

		public string[] listeTypesEquipement(string categorie) {

			string[] tab;
			string str = "SELECT TYPE_EQUIPEMENT,DOMAINE,CATEGORIE FROM TABLE_LISTE_TYPE_EQUIPEMENTS WHERE (CATEGORIE='" + categorie + "')";
			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			DataSet ds = new DataSet();
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter (str, connection1);
			oleDbAdapter1.Fill(ds);
			connection1.Close();

			int nombreItems = ds.Tables[0].Rows.Count;
			tab = new string[nombreItems];

			for(int i = 0; i < nombreItems; i++){
				tab[i] = (string) ds.Tables[0].Rows[i]["TYPE_EQUIPEMENT"];
			}
			return tab;
		}

		public string[] listeCodesOperation(){

			OleDbConnection connection2 = new OleDbConnection(chaineConnection);
			connection2.Open();

			string strSQLOperations = "SELECT CODE_OPERATION,NOM_OPERATION_MAINTENANCE,TYPE_VISITE,UNITE_USAGE,CL_MIN,K_FREQ FROM Liste_OPERATIONS_MPREV WHERE SYSTEMATIQUE='OUI' AND TYPE_EQUIPEMENT ='" + nomType + "'";
			OleDbDataAdapter oleDbAdapter2 = new OleDbDataAdapter (strSQLOperations, connection2);
			DataSet dsCodesOperation = new DataSet();
			oleDbAdapter2.Fill(dsCodesOperation);
			connection2.Close();

			DataTable dt = dsCodesOperation.Tables[0];
			string[] tab = new string[dt.Rows.Count];

			for(int i = 0; i < dt.Rows.Count; i++){
				tab[i] = dt.Rows[i]["CODE_OPERATION"].ToString();
			}
			return tab;
		}

		public string[,] listeOperations(){

			OleDbConnection connection2 = new OleDbConnection(chaineConnection);
			connection2.Open();

			string strSQLOperations = "SELECT CODE_OPERATION,NOM_OPERATION_MAINTENANCE,TYPE_VISITE,UNITE_USAGE,CL_MIN,K_FREQ FROM Liste_OPERATIONS_MPREV WHERE SYSTEMATIQUE='OUI' AND TYPE_EQUIPEMENT ='" + nomType + "'";
			OleDbDataAdapter oleDbAdapter2 = new OleDbDataAdapter (strSQLOperations, connection2);
			DataSet dsCodesOperation = new DataSet();
			oleDbAdapter2.Fill(dsCodesOperation);
			connection2.Close();

			DataTable dt = dsCodesOperation.Tables[0];
			string[,] tab = new string[dt.Rows.Count, 6];

			for(int i = 0; i < dt.Rows.Count; i++){
				tab[i,0] = dt.Rows[i]["CODE_OPERATION"].ToString();
				tab[i,1] = dt.Rows[i]["UNITE_USAGE"].ToString();
				tab[i,2] = dt.Rows[i]["CL_MIN"].ToString();
				tab[i,3] = dt.Rows[i]["K_FREQ"].ToString();
				tab[i,4] = dt.Rows[i]["NOM_OPERATION_MAINTENANCE"].ToString();
				tab[i,5] = dt.Rows[i]["TYPE_VISITE"].ToString();
			}
			return tab;
		}

		public DataSet creerDsOperations(){

			DataSet ds = new DataSet();
			OleDbConnection connection2 = new OleDbConnection(chaineConnection);
			connection2.Open();

			string strSQLOperations = "SELECT CODE_OPERATION,NOM_OPERATION_MAINTENANCE,TYPE_VISITE,CL_MIN,SYSTEMATIQUE FROM Liste_OPERATIONS_MPREV WHERE SYSTEMATIQUE='OUI' AND TYPE_EQUIPEMENT ='" + nomType + "'";
			OleDbDataAdapter oleDbAdapter2 = new OleDbDataAdapter (strSQLOperations, connection2);
			DataSet dsCodesOperation = new DataSet();
			oleDbAdapter2.Fill(ds);
			ds.Tables[0].TableName = "Liste des opérations";
			connection2.Close();
			return ds;
		}

		public string[] chargerCaracteristiquesFMD() {

			string[] tab = new string[20];
			string strSQL = "SELECT NUM3, BETA, ETA, GAMMA, TD_NORM_PAR_HEURE, FMAD, STATUT, N_DEF_PAR_AN, TEMPS_NORM, COUT_NORM_CONS, COUT_NORM_MAT, COUT_NORM_ST, DATE_MODIFICATION, INDICE_FMD FROM TABLE_CARACTERISTIQUES_FMD WHERE (TYPE_EQUIPEMENT= '" + nomType + "')";

			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter(strSQL, connection1);
			connection1.Close();
			DataSet ds = new DataSet();
			oleDbAdapter1.Fill(ds);
			DataTable dt1 = ds.Tables[0];

			if(dt1.Rows.Count != 0){
				tab[0] = double.Parse(dt1.Rows[0]["BETA"].ToString()).ToString();
				tab[1] = double.Parse(dt1.Rows[0]["ETA"].ToString()).ToString("f0");
				tab[2] = double.Parse(dt1.Rows[0]["GAMMA"].ToString()).ToString();
				tab[3] = double.Parse(dt1.Rows[0]["TD_NORM_PAR_HEURE"].ToString()).ToString("#.##E00");
				tab[4] = double.Parse(dt1.Rows[0]["FMAD"].ToString()).ToString("f0");
				tab[5] = dt1.Rows[0]["STATUT"].ToString();
				tab[6] = double.Parse(dt1.Rows[0]["N_DEF_PAR_AN"].ToString()).ToString("f2");
				tab[7] = double.Parse(dt1.Rows[0]["TEMPS_NORM"].ToString()).ToString("f1");
				tab[8] = double.Parse(dt1.Rows[0]["COUT_NORM_MAT"].ToString()).ToString("f0");
				tab[9] = double.Parse(dt1.Rows[0]["COUT_NORM_ST"].ToString()).ToString("f0");
				tab[10] = dt1.Rows[0]["NUM3"].ToString();
				tab[11] = dt1.Rows[0]["DATE_MODIFICATION"].ToString();
				tab[12] = dt1.Rows[0]["INDICE_FMD"].ToString();
				tab[13] = double.Parse(dt1.Rows[0]["COUT_NORM_CONS"].ToString()).ToString("f0");
			}
			else{
				for(int i = 0; i < tab.Length; i++){tab[i] = String.Empty;}
			}
			return tab;
		}

		public double valeurTMRS_MC() {

			double res = 0;
			string strSQL = "SELECT TMRS_MC FROM TABLE_LISTE_OPERATIONS_MCOR WHERE (TYPE_EQUIPEMENT= '" + nomType + "')";

			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter(strSQL, connection1);
			connection1.Close();
			DataSet ds = new DataSet();
			oleDbAdapter1.Fill(ds);
			DataTable dt1 = ds.Tables[0];

			res = double.Parse(dt1.Rows[0]["TMRS_MC"].ToString());
			return res;
		}

		private double calculTD_NORM(){

			double TD_NORM = 0;
			double beta = double.Parse(afficheCarFMD(0));
			double eta = double.Parse(afficheCarFMD(1));
			TD_NORM = beta / eta * Math.Pow((87600 / eta),(beta - 1));
			return TD_NORM;
		}

		private double calculFMAD(){

			double FMAD = 0;
			double beta = double.Parse(afficheCarFMD(0));
			double eta = double.Parse(afficheCarFMD(1));
//			fonction f = new fonction();
//			FMAD =  eta * f.fonctionGamma(1 + 1 / beta);
			return FMAD;
		}

		private double calculN_DEF(){

			double nDef = 8760 * calculTD_NORM();
			return nDef;
		}

		public void calculerCaracteristiquesFMD() {

			modifieCarFMD(3, calculTD_NORM().ToString("#.##E00"));
			modifieCarFMD(4, calculFMAD().ToString("f0"));
			modifieCarFMD(6, calculN_DEF().ToString("f2"));
		}

		public string[] calculValeursCorrigees(double k_cor) {

			string[] tab = new string[6];
			double valeur_TDBASE, valeur_NINT, valeur_ETA;

			string[] tabCarFMD = chargerCaracteristiquesFMD();
			modifieTableauCaracteristiquesFMD(tabCarFMD);

			if(tabCarFMD[0] != String.Empty){
				valeur_TDBASE = double.Parse(tabCarFMD[3]) * k_cor;
				valeur_NINT = double.Parse(tabCarFMD[6]) * k_cor;
				valeur_ETA = double.Parse(tabCarFMD[1]) / k_cor;

				modifieCarFMD(3, valeur_TDBASE.ToString());
				modifieCarFMD(6, valeur_NINT.ToString());
				modifieCarFMD(1, valeur_ETA.ToString("f0"));

				calculerCaracteristiquesFMD();

				tab[0] = tabCarFMD[0];
				tab[1] = tabCarFMD[1];
				tab[2] = tabCarFMD[2];
				tab[3] = tabCarFMD[4];
				tab[4] = tabCarFMD[3];
				tab[5] = tabCarFMD[6];
			}
			else{
				for(int i = 0; i < tab.Length; i++){tab[i] = String.Empty;}
			}
			return tab;
		}

		public void mettreAJourBDD() {

			string strSQL = "SELECT NUM3, BETA, ETA, GAMMA, TD_NORM_PAR_HEURE, FMAD, STATUT, N_DEF_PAR_AN, TEMPS_NORM, COUT_NORM_CONS, COUT_NORM_MAT, COUT_NORM_ST, DATE_MODIFICATION, INDICE_FMD FROM TABLE_CARACTERISTIQUES_FMD WHERE (TYPE_EQUIPEMENT= '" + nomType + "')";

			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter();
			oleDbAdapter1.SelectCommand = new OleDbCommand(strSQL, connection1);
	        OleDbCommandBuilder builder = new OleDbCommandBuilder(oleDbAdapter1);

			connection1.Close();
			DataSet dsCarFMD = new DataSet();
			oleDbAdapter1.Fill(dsCarFMD);

			DataTable dt1 = dsCarFMD.Tables[0];
			modifierLigneTableCarFMD(dt1);
			oleDbAdapter1.Update(dt1);
		}

		private void modifierLigneTableCarFMD(DataTable dt) {

			DataRow dr = dt.Rows[0];

			dr["BETA"] = afficheCarFMD(0);
			dr["ETA"] = afficheCarFMD(1);
			dr["GAMMA"] = afficheCarFMD(2);
			dr["TD_NORM_PAR_HEURE"] = afficheCarFMD(3);
			dr["FMAD"] = afficheCarFMD(4);
			dr["STATUT"] = afficheCarFMD(5);
			dr["N_DEF_PAR_AN"] = afficheCarFMD(6);
			dr["TEMPS_NORM"] = afficheCarFMD(7);
			dr["COUT_NORM_CONS"] = afficheCarFMD(13);
			dr["COUT_NORM_MAT"] = afficheCarFMD(8);
			dr["COUT_NORM_ST"] = afficheCarFMD(9);
			dr["NUM3"] = afficheCarFMD(10);
			dr["DATE_MODIFICATION"] = afficheCarFMD(11);
			dr["INDICE_FMD"] = afficheCarFMD(12);
		}
	}

	//-----------------------------------------------------------------------------------------------------

	public class operation {

		private string code;
		private string type;
		private string[] tableauCaracteristiques;
		private gamme gamme1;
		private string nomBase;


		public void initialiseOperation(string cod, string typ, string[] TabCar, gamme g, string baseName) {

			this.type = typ;
			this.code = cod;
			this.tableauCaracteristiques = TabCar;
			this.gamme1 = g;
			this.nomBase = baseName;
		}

		public string afficheCode{
			get {return code;}
		}

		public string afficheCar(int n){
			return tableauCaracteristiques[n];
		}

		public string[] afficheTabCar{
			get {return tableauCaracteristiques;}
		}

		public gamme afficheGamme{
			get {return gamme1;}
		}

		public void modifieCode(string _code){
			this.code = _code;
		}

		public void modifieGamme(gamme _g){
			this.gamme1 = _g;
		}

		public void modifieCar(int n, string value){
			this.tableauCaracteristiques[n] = value;
		}

		public void modifieTableauCar(string[] tab){
			this.tableauCaracteristiques = tab;
		}

		public void modifieNomBase(string _baseName){
			this.nomBase = _baseName;
		}

		public string[] chargerCaracteristiques() {

			string[] tab = new string[30];
			string chaineConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + "\\" + nomBase + ";Jet OLEDB:Database Password=mexico";

			string strSQL = "SELECT CODE_OPERATION, TYPE_EQUIPEMENT, NOM_OPERATION_MAINTENANCE, TYPE_VISITE, CL_MIN, UNITE_USAGE, SYSTEMATIQUE, CONDITIONNEL, "
				 + "TMRS_MP, K_FREQ FROM TABLE_LISTE_OPERATIONS_MPREV WHERE (CODE_OPERATION= '" + code + "')";

			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter(strSQL, connection1);
			connection1.Close();
			DataSet dsOperation = new DataSet();
			oleDbAdapter1.Fill(dsOperation);

			DataTable dt1 = dsOperation.Tables[0];

			tab[0] = dt1.Rows[0]["TYPE_EQUIPEMENT"].ToString();
			tab[1] = dt1.Rows[0]["NOM_OPERATION_MAINTENANCE"].ToString();
			tab[2] = dt1.Rows[0]["TYPE_VISITE"].ToString();
			tab[3] = dt1.Rows[0]["CL_MIN"].ToString();
			tab[4] = dt1.Rows[0]["UNITE_USAGE"].ToString();
			tab[5] = dt1.Rows[0]["SYSTEMATIQUE"].ToString();
			tab[6] = dt1.Rows[0]["CONDITIONNEL"].ToString();
			tab[7] = dt1.Rows[0]["TMRS_MP"].ToString();
			tab[8] = dt1.Rows[0]["K_FREQ"].ToString();

			return tab;
		}

		public string calculUnitePeriodicite(){

			string res = String.Empty;
			string uniteUsage = afficheCar(4);
			switch (uniteUsage){
			case "TEMPS_CAL" :
				res = "MOIS";
			break;
			case "TEMPS_SERV" :
				res = "HEURE DE SERVICE";
			break;
			case "KM" :
				res = "KM PARCOURU";
			break;
			case "N_MANOEUVRES" :
				res = "MANOEUVRE";
			break;
			case "N_UNITES" :
				res = "UNITE PRODUITE";
			break;
			default :
				res = "MOIS";
			break;
			}
			return res;
		}

		private double calculValeurBase(){

			double res = 0;
			string uniteUsage = afficheCar(4);
			switch (uniteUsage){
			case "TEMPS_CAL" :
				res = 12;
			break;
			case "TEMPS_SERV" :
				res = 8000;
			break;
			case "KM" :
				res = 15000;
			break;
			case "N_MANOEUVRES" :
				res = 10000;
			break;
			case "N_UNITES" :
				res = 10000;
			break;
			default :
				res = 12;
			break;
			}
			return res;
		}

		public double calculPeriodicite(double _cl_G, double K_CL_G_0, double K_CL_G_1, double K_CL_G_2){

			double res = 0;
			double valeurBase = calculValeurBase();
			double kFreq = double.Parse(afficheCar(8));
			double classeMin = double.Parse(afficheCar(3));

			switch (_cl_G.ToString()){
			case "0" :
				if(valeurBase > 1000){
					res = (Math.Round((valeurBase / K_CL_G_0 / kFreq)/100,0)*100);
				}
				else{
					double valeur0 = (valeurBase / K_CL_G_0 / kFreq);
					if(valeur0 < 1){valeur0 = 1;}
					res = Math.Round(valeur0,0);
				}
			break;
			case "1" :
				double K_CL_G_a = 0;
				if(_cl_G > classeMin){K_CL_G_a = K_CL_G_1;}else{K_CL_G_a = 1.00;}
				if(valeurBase > 1000){
					res = (Math.Round((valeurBase / K_CL_G_a / kFreq)/100,0)*100);
				}
				else{
					double valeur1 = (valeurBase / K_CL_G_a / kFreq);
					if(valeur1 < 1){valeur1 = 1;}
					res = Math.Round(valeur1,0);
				}
			break;
			case "2" :
				double K_CL_G_b = 0;
				if(_cl_G > classeMin){K_CL_G_b = K_CL_G_2;}else{K_CL_G_b = 1.00;}
				if(valeurBase > 1000){
					res = (Math.Round((valeurBase / K_CL_G_b / kFreq)/100,0)*100);
				}
				else{
					double valeur2 = (valeurBase / K_CL_G_b / kFreq);
					if(valeur2 < 1){valeur2 = 1;}
					res = Math.Round(valeur2,0);
				}
			break;
			}
			return res;
		}
	}

	//-----------------------------------------------------------------------------------------------------

	public class gamme {

		private string code;
		private string[] tableauOperationsElementaires;
		private string nomBase = String.Empty;
		private string chaineConnection = String.Empty;


		public void initialiseGamme(string baseName, string cod, string[] tabOpElem) {

			this.nomBase = baseName;
			this.code = cod;
			this.tableauOperationsElementaires = tabOpElem;
			this.chaineConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + "\\" + nomBase + ";Jet OLEDB:Database Password=mexico";

		}

		public string afficheCode{
			get {return code;}
		}

		public string afficheOpElem(int n){
			return tableauOperationsElementaires[n];
		}

		public string[] afficheTabOpElem{
			get {return tableauOperationsElementaires;}
		}

		public void modifieCode(string _code){
			this.code = _code;
		}

		public void modifieOpElem(int n, string value){
			this.tableauOperationsElementaires[n] = value;
		}

		public void modifieTableauOpElem(string[] tab){
			this.tableauOperationsElementaires = tab;
		}

		public bool testExist(){
			bool res = false;
			string strSQL = "SELECT CODE_OPERATION FROM TABLE_LISTE_GAMMES_OPERATION WHERE (CODE_OPERATION= '" + code + "')";
			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter(strSQL, connection1);
			connection1.Close();
			DataSet ds = new DataSet();
			oleDbAdapter1.Fill(ds);
			if(ds.Tables[0].Rows.Count != 0){res = true;}else{res = false;}
			return res;
		}

		public bool testExist(DataTable dt){
			bool res = false;
			string selection = "CODE_OPERATION ='" + code + "'";
			DataRow[] tabdr = dt.Select(selection);
			if(tabdr.Length != 0){
				res = true;
			}
			else{
				res = false;
			}
			return res;
		}

		public string[] chargerGamme() {

			string[] tab = new string[80];
			string strSQL = "SELECT NUMERO, INDICE_GAMME, CODE_OPERATION, TYPE_EQUIPEMENT, CONSIGNE, NON_CONSIGNE, OPERATION_1, OPERATION_2, OPERATION_3, OPERATION_4, OPERATION_5, OPERATION_6,"
			 + " OPERATION_7, OPERATION_8, OPERATION_9, OPERATION_10, OPERATION_11, OPERATION_12, OPERATION_13, OPERATION_14, OPERATION_15, OPERATION_16, OPERATION_17,"
			 + " OPERATION_18, OPERATION_19, OPERATION_20, TEMPS_NORM, COUT_NORM_CONS, COUT_NORM_PIECES, COUT_NORM_ST, LISTE_OUTILLAGE, LISTE_CONSOMMABLES, Lien_Documentation, LISTE_autres_documents, HABILITATIONS, QUALIFICATIONS, EFF_MOY, LISTE_COMPETENCES_PARTICULIERES FROM TABLE_LISTE_GAMMES_OPERATION WHERE (CODE_OPERATION= '" + code + "')";

			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter(strSQL, connection1);
			connection1.Close();
			DataSet dsOperation = new DataSet();
			oleDbAdapter1.Fill(dsOperation);

			DataTable dt1 = dsOperation.Tables[0];

			tab[0] = dt1.Rows[0]["NUMERO"].ToString();
			tab[1] = dt1.Rows[0]["INDICE_GAMME"].ToString();
			tab[2] = dt1.Rows[0]["CONSIGNE"].ToString();
			tab[3] = dt1.Rows[0]["NON_CONSIGNE"].ToString();
			tab[10] = dt1.Rows[0]["OPERATION_1"].ToString();
			tab[11] = dt1.Rows[0]["OPERATION_2"].ToString();
			tab[12] = dt1.Rows[0]["OPERATION_3"].ToString();
			tab[13] = dt1.Rows[0]["OPERATION_4"].ToString();
			tab[14] = dt1.Rows[0]["OPERATION_5"].ToString();
			tab[15] = dt1.Rows[0]["OPERATION_6"].ToString();
			tab[16] = dt1.Rows[0]["OPERATION_7"].ToString();
			tab[17] = dt1.Rows[0]["OPERATION_8"].ToString();
			tab[18] = dt1.Rows[0]["OPERATION_9"].ToString();
			tab[19] = dt1.Rows[0]["OPERATION_10"].ToString();
			tab[20] = dt1.Rows[0]["OPERATION_11"].ToString();
			tab[21] = dt1.Rows[0]["OPERATION_12"].ToString();
			tab[22] = dt1.Rows[0]["OPERATION_13"].ToString();
			tab[23] = dt1.Rows[0]["OPERATION_14"].ToString();
			tab[24] = dt1.Rows[0]["OPERATION_15"].ToString();
			tab[25] = dt1.Rows[0]["OPERATION_16"].ToString();
			tab[26] = dt1.Rows[0]["OPERATION_17"].ToString();
			tab[27] = dt1.Rows[0]["OPERATION_18"].ToString();
			tab[28] = dt1.Rows[0]["OPERATION_19"].ToString();
			tab[29] = dt1.Rows[0]["OPERATION_20"].ToString();
			tab[30] = dt1.Rows[0]["TEMPS_NORM"].ToString();
			tab[31] = dt1.Rows[0]["COUT_NORM_CONS"].ToString();
			tab[32] = dt1.Rows[0]["COUT_NORM_ST"].ToString();
			tab[33] = dt1.Rows[0]["LISTE_OUTILLAGE"].ToString();
			tab[34] = dt1.Rows[0]["LISTE_CONSOMMABLES"].ToString();
			tab[35] = dt1.Rows[0]["Lien_Documentation"].ToString();
			tab[36] = dt1.Rows[0]["LISTE_autres_documents"].ToString();
			tab[37] = dt1.Rows[0]["HABILITATIONS"].ToString();
			tab[38] = dt1.Rows[0]["QUALIFICATIONS"].ToString();
			tab[39] = dt1.Rows[0]["EFF_MOY"].ToString();
			tab[40] = dt1.Rows[0]["LISTE_COMPETENCES_PARTICULIERES"].ToString();
			tab[41] = dt1.Rows[0]["COUT_NORM_PIECES"].ToString();

			return tab;
		}

		public void mettreAJourBDD() {

			string strSQL = "SELECT NUMERO, INDICE_GAMME, CODE_OPERATION, TYPE_EQUIPEMENT, CONSIGNE, NON_CONSIGNE, OPERATION_1, OPERATION_2, OPERATION_3, OPERATION_4, OPERATION_5, OPERATION_6,"
			 + " OPERATION_7, OPERATION_8, OPERATION_9, OPERATION_10, OPERATION_11, OPERATION_12, OPERATION_13, OPERATION_14, OPERATION_15, OPERATION_16, OPERATION_17,"
			 + " OPERATION_18, OPERATION_19, OPERATION_20, TEMPS_NORM, COUT_NORM_CONS, COUT_NORM_ST, LISTE_OUTILLAGE, LISTE_CONSOMMABLES, Lien_Documentation, LISTE_autres_documents, HABILITATIONS, QUALIFICATIONS, EFF_MOY, LISTE_COMPETENCES_PARTICULIERES FROM TABLE_LISTE_GAMMES_OPERATION WHERE (CODE_OPERATION= '" + code + "')";

			OleDbConnection connection1 = new OleDbConnection(chaineConnection);
			connection1.Open();
			OleDbDataAdapter oleDbAdapter1 = new OleDbDataAdapter();
			oleDbAdapter1.SelectCommand = new OleDbCommand(strSQL, connection1);
	        OleDbCommandBuilder builder = new OleDbCommandBuilder(oleDbAdapter1);

			connection1.Close();
			DataSet dsOperation = new DataSet();
			oleDbAdapter1.Fill(dsOperation);

			DataTable dt1 = dsOperation.Tables[0];
			modifierLigneTableGamme(dt1);
			oleDbAdapter1.Update(dt1);
		}

		private void modifierLigneTableGamme(DataTable dt) {

			DataRow dr = dt.Rows[0];

			dr["NUMERO"] = afficheOpElem(0);
			// tab[0] = dt1.Rows[0]["NUMERO"].ToString();
			dr["INDICE_GAMME"] = afficheOpElem(1);
			dr["CONSIGNE"] = afficheOpElem(2);
			dr["NON_CONSIGNE"] = afficheOpElem(3);
			dr["OPERATION_1"] = afficheOpElem(10);
			dr["OPERATION_2"] = afficheOpElem(11);
			dr["OPERATION_3"] = afficheOpElem(12);
			dr["OPERATION_4"] = afficheOpElem(13);
			dr["OPERATION_5"] = afficheOpElem(14);
			dr["OPERATION_6"] = afficheOpElem(15);
			dr["OPERATION_7"] = afficheOpElem(16);
			dr["OPERATION_8"] = afficheOpElem(17);
			dr["OPERATION_9"] = afficheOpElem(18);
			dr["OPERATION_10"] = afficheOpElem(19);
			dr["OPERATION_11"] = afficheOpElem(20);
			dr["OPERATION_12"] = afficheOpElem(21);
			dr["OPERATION_13"] = afficheOpElem(22);
			dr["OPERATION_14"] = afficheOpElem(23);
			dr["OPERATION_15"] = afficheOpElem(24);
			dr["OPERATION_16"] = afficheOpElem(25);
			dr["OPERATION_17"] = afficheOpElem(26);
			dr["OPERATION_18"] = afficheOpElem(27);
			dr["OPERATION_19"] = afficheOpElem(28);
			dr["OPERATION_20"] = afficheOpElem(29);
			dr["TEMPS_NORM"] = double.Parse(afficheOpElem(30));
			dr["COUT_NORM_CONS"] = double.Parse(afficheOpElem(31));
			dr["COUT_NORM_ST"] = double.Parse(afficheOpElem(32));
			dr["LISTE_OUTILLAGE"] = afficheOpElem(33);
			dr["LISTE_CONSOMMABLES"] = afficheOpElem(34);
			dr["Lien_Documentation"] = afficheOpElem(35);
			dr["LISTE_autres_documents"] = afficheOpElem(36);
			dr["HABILITATIONS"] = afficheOpElem(37);
			dr["QUALIFICATIONS"] = afficheOpElem(38);
			dr["EFF_MOY"] = double.Parse(afficheOpElem(39));
			dr["LISTE_COMPETENCES_PARTICULIERES"] = afficheOpElem(40);
		}
	}

	//-----------------------------------------------------------------------------------------------------

	public class visu_base{

		private string[] dom;
		private string[] cat;
		private string chaineSQL = String.Empty;


		public visu_base(string[] domaines, string[] categories){

			this.dom = domaines;
			this.cat = categories;
		}

		public string[] domaines{
			get {return dom;}
		}

		public void modifieDomaines(string[] _dom){
			this.dom = _dom;
		}

		public string creerChaineSQL_TypeEquipements(){
			string res = String.Empty;
			string domaines = selectionDomaines();
			res =  "SELECT ESPACE,DOMAINE,CATEGORIE,TYPE,SOUSTYPE_1,SOUSTYPE_2,SOUSTYPE_3,TYPE_EQUIPEMENT FROM TABLE_LISTE_TYPE_EQUIPEMENTS WHERE (DOMAINE = '" + domaines + " ORDER BY TYPE_EQUIPEMENT";
			return res;
		}

		public string creerChaineSQL_PREV(){
			string res = String.Empty;
			string domaines = selectionDomaines();
			res =  "SELECT TYPE_EQUIPEMENT,CODE_OPERATION,NOM_OPERATION_MAINTENANCE,TYPE_VISITE,CL_MIN,UNITE_USAGE,SYSTEMATIQUE,CONDITIONNEL,TMRS_MP FROM Liste_OPERATIONS_MPREV WHERE  (DOMAINE='" + domaines + " ORDER BY TYPE_EQUIPEMENT";
			return res;
		}

		public string creerChaineSQL_CORR(){
			string res = String.Empty;
			string domaines = selectionDomaines();
			res =  "SELECT TYPE_EQUIPEMENT,CODE_OPERATION,NOM_OPERATION_MAINTENANCE,TMRS_MC,N_DEF_PAR_AN FROM Liste_Operations_MCOR WHERE (DOMAINE='" + domaines + " ORDER BY TYPE_EQUIPEMENT";
			return res;
		}

		public string creerChaineSQL_FMD(){
			string res = String.Empty;
			string domaines = selectionDomaines();
			res =  "SELECT DOMAINE,CATEGORIE,TYPE_EQUIPEMENT,BETA,ETA,GAMMA,TD_NORM_PAR_HEURE,N_DEF_PAR_AN,FMAD,STATUT FROM Liste_CARACTERISTIQUES_FMD WHERE  (DOMAINE='" + domaines +  " ORDER BY TYPE_EQUIPEMENT" ;
			return res;
		}

		public string afficheChaineSQL{
			get {return chaineSQL;}
		}

		public bool testListeDom(){
			bool res = false;
			foreach(string item in dom){
				if(item != String.Empty){
					bool temp = true;
					res = res | temp;
				}
			}
			return res;
		}

		private string selectionDomaines(){
			string res = String.Empty;
			for(int i = 0; i < dom.Length; i++){
				if(res == ""){
					res = dom[i] + "')";
				}
				else{
					res = res + " OR (DOMAINE='" + dom[i] + "')";
				}
			}
			return res;
		}
	}

	//-----------------------------------------------------------------------------------------------------

	public class bilan {


		private projet projet1;
		private liste_equipements liste1;
		private liste_plan liste2;
		private string[] _systemes;
		private string[] _secteurs;

		private Array _equipements_par_systemes_total; // unidimensionnel - col :  classeG
		private Array _equipements_par_systemes_perim; // unidimensionnel - col :  classeG
		private Array _equipements_par_secteurs_total; // unidimensionnel - col :  classeG
		private Array _equipements_par_secteurs_perim; // unidimensionnel - col :  classeG
		private Array _charge_MO_MP_par_systemes; // unidimensionnel - col :  classeG
		private Array _charge_CONS_MP_par_systemes; // unidimensionnel - col :  classeG
		private Array _charge_MO_MC_par_systemes; // unidimensionnel - col :  classeG

		private DataTable dt1;


		public bilan(projet p) {

			this.projet1 = p;
			this.liste1 = projet1.listeEquipements();
			this.liste2 = projet1.planMaintenance();

			initialisation();
		}

		private void initialisation() {

			this._systemes = projet1.listeSystemes();
			this._secteurs = projet1.listeSecteurs_complet();

			_equipements_par_systemes_total = Array.CreateInstance( typeof(double[]), 3 );
			_equipements_par_systemes_perim = Array.CreateInstance( typeof(double[]), 3 );
			_equipements_par_secteurs_total = Array.CreateInstance( typeof(double[]), 3 );
			_equipements_par_secteurs_perim = Array.CreateInstance( typeof(double[]), 3 );
			_charge_MO_MP_par_systemes = Array.CreateInstance( typeof(double[]), 4 ); // col 4 = somme des 3 premières
			_charge_CONS_MP_par_systemes = Array.CreateInstance( typeof(double[]), 4 ); // col 4 = somme des 3 premières
			_charge_MO_MC_par_systemes = Array.CreateInstance( typeof(double[]), 4 ); // col 4 = somme des 3 premières
		}

		public string nom{
			get {return projet1.afficheNom;}
		}

		public string[] systemes{
			get {return _systemes;}
		}

		public string[] secteurs{
			get {return _secteurs;}
		}

		public Array equipements_par_systemes_total{
			get{return _equipements_par_systemes_total;}
		}

		public Array equipements_par_systemes_perim{
			get{return _equipements_par_systemes_perim;}
		}

		public Array equipements_par_secteurs_total{
			get{return _equipements_par_secteurs_total;}
		}

		public Array equipements_par_secteurs_perim{
			get{return _equipements_par_secteurs_perim;}
		}

		public Array charge_MO_MP_par_systemes{
			get{return _charge_MO_MP_par_systemes;}
		}

		public Array charge_MO_MC_par_systemes{
			get{return _charge_MO_MC_par_systemes;}
		}

		public Array charge_CONS_MP_par_systemes{
			get{return _charge_CONS_MP_par_systemes;}
		}

		private double totalCharge_MP(string champ) {
			double res = 0;

			return res;
		}

		private double totalCharge_MC(string champ) {
			double res = 0;

			return res;
		}

		public void creerChampsSystemes(DataTable dt){
			bool[] tab = new bool[]{false, true};
			string champ = String.Empty;
			foreach(bool b in tab){
				if(b == false){champ = "TOTAL";}else{champ = "PERIM";}
				foreach(string item in _systemes){
					DataRow dr = dt.NewRow();
					dr["REPERE"] = "SYST_" + champ;
					dr["DESIGNATION"] = item;
					dr["VALEUR_0"] = liste1.nombreEquipementsParSysteme(0, item, b);
					dr["VALEUR_1"] = liste1.nombreEquipementsParSysteme(1, item, b);
					dr["VALEUR_2"] = liste1.nombreEquipementsParSysteme(2, item, b);
					dt.Rows.Add(dr);
				}
			}
		}

		public void creerChampsSecteurs(DataTable dt){
			bool[] tab = new bool[]{false, true};
			string champ = String.Empty;
			foreach(bool b in tab){
				if(b == false){champ = "TOTAL";}else{champ = "PERIM";}
				foreach(string item in _secteurs){
					DataRow dr = dt.NewRow();
					dr["REPERE"] = "SECT_" + champ;
					dr["DESIGNATION"] = item;
					dr["VALEUR_0"] = liste1.nombreEquipementsParSecteur(0, item, b);
					dr["VALEUR_1"] = liste1.nombreEquipementsParSecteur(1, item, b);
					dr["VALEUR_2"] = liste1.nombreEquipementsParSecteur(2, item, b);
					dt.Rows.Add(dr);
				}
			}
		}

		public void creerChampsCharge_MP(DataTable dt){

			string[] tab = new string[]{"MO", "CONS", "PR", "ST"};
			foreach(string champ in tab){
				foreach(string item in _systemes){
					DataRow dr = dt.NewRow();
					dr["REPERE"] = "CHARGE_" + champ + "_PREV_SYST";
					dr["DESIGNATION"] = item;
					dr["VALEUR_0"] = Math.Round(projet1.charge_MPREV(item, 0, champ), 1);
					dr["VALEUR_1"] = Math.Round(projet1.charge_MPREV(item, 1, champ), 1);
					dr["VALEUR_2"] = Math.Round(projet1.charge_MPREV(item, 2, champ), 1);
					dt.Rows.Add(dr);
				}
			}
		}

		public void creerChampsCharge_MC(DataTable dt){

			string[] tab = new string[]{"MO", "CONS", "PR", "ST"};
			foreach(string champ in tab){
				foreach(string item in _systemes){
					DataRow dr = dt.NewRow();
					dr["REPERE"] = "CHARGE_" + champ + "_COR_SYST";
					dr["DESIGNATION"] = item;
					dr["VALEUR_0"] = Math.Round(projet1.charge_MCOR(item, 0, champ), 1);
					dr["VALEUR_1"] = Math.Round(projet1.charge_MCOR(item, 1, champ), 1);
					dr["VALEUR_2"] = Math.Round(projet1.charge_MCOR(item, 2, champ), 1);
					dt.Rows.Add(dr);
				}
			}
		}

		public void configurerEquipements_systemes(DataTable dt){

			string[] tab = new string[]{"TOTAL", "PERIM"};
			string selection = String.Empty;
			foreach(string item in tab){
				selection = "REPERE ='SYST_" + item + "'";
				DataRow[] dr = dt.Select(selection);
				double[] tab0 = new double[dr.Length];
				double[] tab1 = new double[dr.Length];
				double[] tab2 = new double[dr.Length];

				for(int i = 0; i < dr.Length; i++){
					tab0[i] = double.Parse(dr[i]["VALEUR_0"].ToString());
					tab1[i] = double.Parse(dr[i]["VALEUR_1"].ToString());
					tab2[i] = double.Parse(dr[i]["VALEUR_2"].ToString());
					switch(item){
					case "TOTAL" :
						_equipements_par_systemes_total.SetValue(tab0, 0);
						_equipements_par_systemes_total.SetValue(tab1, 1);
						_equipements_par_systemes_total.SetValue(tab2, 2);
					break;
					case "PERIM" :
						_equipements_par_systemes_perim.SetValue(tab0, 0);
						_equipements_par_systemes_perim.SetValue(tab1, 1);
						_equipements_par_systemes_perim.SetValue(tab2, 2);
					break;
					}
				}
			}
		}

		public void configurerEquipements_secteurs(DataTable dt){

			string[] tab = new string[]{"TOTAL", "PERIM"};
			string selection = String.Empty;
			foreach(string item in tab){
				selection = "REPERE ='SECT_" + item + "'";
				DataRow[] dr = dt.Select(selection);
				double[] tab0 = new double[dr.Length];
				double[] tab1 = new double[dr.Length];
				double[] tab2 = new double[dr.Length];

				for(int i = 0; i < dr.Length; i++){
					tab0[i] = double.Parse(dr[i]["VALEUR_0"].ToString());
					tab1[i] = double.Parse(dr[i]["VALEUR_1"].ToString());
					tab2[i] = double.Parse(dr[i]["VALEUR_2"].ToString());
					switch(item){
					case "TOTAL" :
						_equipements_par_secteurs_total.SetValue(tab0, 0);
						_equipements_par_secteurs_total.SetValue(tab1, 1);
						_equipements_par_secteurs_total.SetValue(tab2, 2);
					break;
					case "PERIM" :
						_equipements_par_secteurs_perim.SetValue(tab0, 0);
						_equipements_par_secteurs_perim.SetValue(tab1, 1);
						_equipements_par_secteurs_perim.SetValue(tab2, 2);
					break;
					}
				}
			}
		}

		public void configurerCharge_MO_MP_systemes(DataTable dt){

			string selection = String.Empty;
			selection = "REPERE ='CHARGE_MO_PREV_SYST'";
			DataRow[] dr = dt.Select(selection);
			double[] tab0 = new double[dr.Length];
			double[] tab1 = new double[dr.Length];
			double[] tab2 = new double[dr.Length];
			double[] tab3 = new double[dr.Length]; // somme

			for(int i = 0; i < dr.Length; i++){
				tab0[i] = double.Parse(dr[i]["VALEUR_0"].ToString());
				tab1[i] = double.Parse(dr[i]["VALEUR_1"].ToString());
				tab2[i] = double.Parse(dr[i]["VALEUR_2"].ToString());
				tab3[i] = tab0[i] + tab1[i] + tab2[i];
			}
			_charge_MO_MP_par_systemes.SetValue(tab0, 0);
			_charge_MO_MP_par_systemes.SetValue(tab1, 1);
			_charge_MO_MP_par_systemes.SetValue(tab2, 2);
			_charge_MO_MP_par_systemes.SetValue(tab3, 3);
		}

		public void configurerCharge_MO_MC_systemes(DataTable dt){

			string selection = String.Empty;
			selection = "REPERE ='CHARGE_MO_COR_SYST'";
			DataRow[] dr = dt.Select(selection);
			double[] tab0 = new double[dr.Length];
			double[] tab1 = new double[dr.Length];
			double[] tab2 = new double[dr.Length];
			double[] tab3 = new double[dr.Length]; // somme

			for(int i = 0; i < dr.Length; i++){
				tab0[i] = double.Parse(dr[i]["VALEUR_0"].ToString());
				tab1[i] = double.Parse(dr[i]["VALEUR_1"].ToString());
				tab2[i] = double.Parse(dr[i]["VALEUR_2"].ToString());
				tab3[i] = tab0[i] + tab1[i] + tab2[i];
			}
			_charge_MO_MC_par_systemes.SetValue(tab0, 0);
			_charge_MO_MC_par_systemes.SetValue(tab1, 1);
			_charge_MO_MC_par_systemes.SetValue(tab2, 2);
			_charge_MO_MC_par_systemes.SetValue(tab3, 3);
		}

		public void chargerBilan(){

			dt1 = projet1.afficheDsProjet.Tables[7];
			configurerEquipements_systemes(dt1);
			configurerEquipements_secteurs(dt1);
			configurerCharge_MO_MP_systemes(dt1);
			configurerCharge_MO_MC_systemes(dt1);
		}

	}

	//-----------------------------------------------------------------------------------------------------
}
