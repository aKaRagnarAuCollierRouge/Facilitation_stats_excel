import sys
import ast
import openpyxl
from PySide2.QtGui import QIcon
from PySide2.QtWidgets import QApplication, QMainWindow, QFileDialog, QVBoxLayout, QWidget, QPushButton, QLabel, \
    QCheckBox, QRadioButton, QComboBox, QScrollArea, QLineEdit, QGridLayout, QTextEdit, QTabWidget, QAction
from PySide2.QtCore import Qt
import json
from Style import *
from Fenetre_modifier_data_default.fenetre_value_default_excel import Change_value_filtre_tri_excel_window
from Fonctions_complementaires import *
class FilePickerApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.initUI()
        # Listes des widgets créers
        self.liste_widgets_critères = []
        self.liste_widgets_datas_extraire = []

    def initUI(self):
        self.setWindowTitle("Gear5:Traitement datas et analyse")
        self.setGeometry(100, 100, 400, 200)
        self.setWindowIcon(QIcon("C:\\Users\\Baptiste\\Documents\\GUI_ext_tri_datas\\Gear5.jpg"))


        # Création d'une barre d'outils
        self.menuBar = self.menuBar()
        self.default_menu=self.menuBar.addMenu("&Changer valeurs defaults")
        self.change_excel_default_value = QAction("Changer valeurs_default")
        self.default_menu.addAction(self.change_excel_default_value)
        self.change_excel_default_value.triggered.connect(self.fct_change_value_trie_filtrage)



        # Création table et différents onglets
        self.tabs = QTabWidget()
        self.tabs.setTabPosition(QTabWidget.East)
        self.setCentralWidget(self.tabs)
        self.onglet_trie_filtrage = QWidget()
        self.onglet_stats= QWidget()
        self.transfert_colonne=QWidget()
        self.tabs.addTab(self.onglet_trie_filtrage,"Trie Filtrage Exel")
        self.tabs.addTab(self.onglet_stats,"Statistiques")
        self.tabs.addTab(self.transfert_colonne,"Colonne F to F")
        layout=QGridLayout()
        layout_stats=QGridLayout()
        layout_transfert_F_to_F=QGridLayout()
        self.onglet_trie_filtrage.setLayout(layout)
        self.onglet_stats.setLayout(layout_stats)
        self.transfert_colonne.setLayout(layout_transfert_F_to_F)



        #  QPUSH SELECTION FICHIER
        self.select_file_button = QPushButton("Sélectionner un fichier", self)



        self.select_file_button.clicked.connect(self.showFileDialog)

         # QLabel AFFICHAGE CHEMIN
        self.file_path_label = QLineEdit("", self)
        self.file_path_label.setReadOnly(True)



        # QPushButton APPLIQUER QU'ON A BIEN SELECTION

        self.Button_extraction_feuilles_critères=QPushButton("Confirmer la Selection")
        self.Button_extraction_feuilles_critères.clicked.connect(self.extraction_critères)


        #Ajouter critères de selection
        self.Label_ajout_critère=QLabel("Ajouter des critères de traitement du fichier:")
        self.Cb_feuille_critère=QComboBox()
        self.Cb_feuille_critère.currentTextChanged.connect(self.change_combobox_Feuille)
        self.Cb_critères=QComboBox()
        self.Opérateurs=QComboBox()
        liste_opérateurs=["=","<",">","<=",">=","contient","ne contient pas"]
        self.Opérateurs.addItems(liste_opérateurs)
        self.Comparaison=QLineEdit()
        self.Btn_ajouter_critères=QPushButton("+Ajouter le critère de  selection+")
        self.Btn_ajouter_critères.clicked.connect(self.Ajouter_critères)
        self.Btn_remove_last_critère=QPushButton("-Effacer dernier critère-")
        self.Btn_remove_last_critère.clicked.connect(self.remove_last_widget_scroll_area)




        # Créer une zone de défilement pour les CRITERES
        scroll_area = QScrollArea(self)
        widget_container = QWidget(self)
        scroll_area.setWidget(widget_container)
        scroll_area.setWidgetResizable(True)

        self.scroll_layout = QGridLayout(widget_container)


        #Créer une zone de défilement pour les datas à Extraire avec Bouton +
        self.Label_extraire=QLabel("Données à Extraire des fichiers(colonne de la ligne):")
        self.Cb_feuille_extraire=QComboBox()
        self.Cb_feuille_extraire.currentTextChanged.connect(self.change_combobox_Feuille)
        self.Cb_colonne_extraire=QComboBox()
        self.Btn_data_extraire = QPushButton("+Données à extraire+")

        self.Btn_data_remove_extraire=QPushButton("-Effacer derniere selection-")
        self.Btn_data_remove_extraire.clicked.connect(self.remove_last_widget_scroll_area_extraire)

        self.Btn_data_extraire.clicked.connect(self.Ajouter_data_extraire)


        scroll_area_extraire = QScrollArea(self)
        widget_container_extraire = QWidget(self)
        scroll_area_extraire.setWidget(widget_container_extraire)
        scroll_area_extraire.setWidgetResizable(True)

        self.scroll_layout_extraire = QVBoxLayout(widget_container_extraire)


        self.label_name_sheet_excel = QLabel("Veuillez indiquer le nom de la feuille excel:")
        self.Name_sheet = QLineEdit()
        self.Btn_Ajouter_feuille=QPushButton("+Ajouter Feuille+")
        self.Btn_Ajouter_feuille.setStyleSheet(style_btn_ajouter)
        self.Btn_Ajouter_feuille.clicked.connect(self.Ajouter_feuille)
        self.Btn_Supprimer_feuille=QPushButton("-Supprimer Feuille-")
        self.Btn_Supprimer_feuille.setStyleSheet(style_btn_supprimer)
        self.Btn_Supprimer_feuille.clicked.connect(self.Supprimer_feuille)
        self.label_feuilles_created = QLabel("Feuilles à créer:")
        scroll_area_feuilles_created = QScrollArea(self)
        widget_container_feuilles_created = QWidget(self)
        scroll_area_feuilles_created.setWidget(widget_container_feuilles_created)
        scroll_area_feuilles_created.setWidgetResizable(True)
        self.scroll_layout_feuilles_created = QVBoxLayout(widget_container_feuilles_created)

        self.label_name_fichier_excel=QLabel("Quel est le nom du fichier excel:")
        self.name_fichier_excel=QLineEdit()

        #QCheckBox pour savoir si on applique data dans nouveau excel ou un autre spécifique
        self.button_select_folder=QPushButton("Selectionner dossier où écrire le excel")
        self.button_select_folder.clicked.connect(self.select_folder)
        self.selection_result_folder=QLineEdit()
        self.selection_result_folder.setReadOnly(True)

        self.button_appliquer_to_excel=QPushButton("Trier données et exporter vers excel")

        self.button_appliquer_to_excel.clicked.connect(self.Traitement_données)

        layout.addWidget(self.select_file_button)
        layout.addWidget(self.file_path_label)
        layout.addWidget(self.Button_extraction_feuilles_critères)
        layout.addWidget(self.Label_ajout_critère)
        layout.addWidget(self.Cb_feuille_critère)
        layout.addWidget(self.Cb_critères)
        layout.addWidget(self.Opérateurs)
        layout.addWidget(self.Comparaison)
        layout.addWidget(self.Btn_ajouter_critères)
        layout.addWidget(self.Btn_remove_last_critère)
        layout.addWidget(scroll_area)
        layout.addWidget(self.Label_extraire)
        layout.addWidget(self.Cb_feuille_extraire)
        layout.addWidget(self.Cb_colonne_extraire)
        layout.addWidget(self.Btn_data_extraire)
        layout.addWidget(self.Btn_data_remove_extraire)
        layout.addWidget(scroll_area_extraire)
        layout.addWidget(self.label_name_sheet_excel)
        layout.addWidget(self.Name_sheet)
        layout.addWidget(self.Btn_Ajouter_feuille)
        layout.addWidget(self.Btn_Supprimer_feuille)
        layout.addWidget(self.label_feuilles_created)
        layout.addWidget(scroll_area_feuilles_created)
        layout.addWidget(self.label_name_fichier_excel)
        layout.addWidget(self.name_fichier_excel)
        layout.addWidget(self.button_select_folder)
        layout.addWidget(self.selection_result_folder)
        layout.addWidget(self.button_appliquer_to_excel)


#----------------------------------WIDGET 2EME ONGLET--------------------------------------------------------
        self.radio_stats_ref=QRadioButton("Stat with coll ref")
        self.radio_stats_ref.setChecked(True)
        self.radio_stats_ref.toggled.connect(self.change_widget_radio_stats)
        self.radio_stats_without_ref=QRadioButton("Stat without coll ref")
        self.radio_stats_without_ref.toggled.connect(self.change_widget_radio_stats)
        self.select_fichier_stat=QPushButton("Selectionner Fichier")
        self.select_fichier_stat.clicked.connect(self.showFileDialog)
        self.Confirm_select_fichier_stat=QPushButton("Confirme la selection du fichier")
        self.Confirm_select_fichier_stat.clicked.connect(self.extraction_critères)
        self.chemin_fichier_stat=QLineEdit()
        self.chemin_fichier_stat.setReadOnly(True)
        self.label_select_col_ref=QLabel("Selection de la colonne de référence:")
        self.Cb_feuille_ref=QComboBox()
        self.Cb_Colonne_ref=QComboBox()
        self.Cb_feuille_ref.currentTextChanged.connect(self.change_combobox_Feuille)
        self.label_select_colonnes_secondaires=QLabel("Selection des colonnes secondaires:")
        self.Cb_feuille_coll_secondaire=QComboBox()
        self.Cb_coll_secondaire=QComboBox()
        self.Cb_feuille_coll_secondaire.currentTextChanged.connect(self.change_combobox_Feuille)
        self.Button_ajouter_coll_secondaire=QPushButton("+Ajouter colonne secondaire+")
        self.Button_ajouter_coll_secondaire.clicked.connect(self.Ajouter_colonne_secondaires)
        self.Button_supprimer_coll_secondaire=QPushButton("-Supprimer colonne secondaire-")
        self.Button_supprimer_coll_secondaire.clicked.connect(self.remove_last_widget_scroll_area_coll_secondaires)

        #Creation de la Scroll Area
        Scroll_Area_coll_secondaire=QScrollArea(self)
        widget_container_coll_secondaire = QWidget(self)
        Scroll_Area_coll_secondaire.setWidget(widget_container_coll_secondaire)
        Scroll_Area_coll_secondaire.setWidgetResizable(True)
        self.scroll_layout_coll_secondaire = QGridLayout(widget_container_coll_secondaire)


        self.coched_exclure_datas=QCheckBox("Exclure les cases vides et autres datas spécifié(word_exclure_df)")
        self.label_nom_feuille_created=QLabel("Voulez vous donnez un nom spécifique à la feuille créer?")
        self.nom_feuille_created=QLineEdit()
        self.Button_appliquer=QPushButton("Appliquer")
        self.Button_appliquer.clicked.connect(self.Construire_appliquer_stats)

        layout_stats.addWidget(self.radio_stats_ref)
        layout_stats.addWidget(self.radio_stats_without_ref)
        layout_stats.addWidget(self.select_fichier_stat)
        layout_stats.addWidget(self.chemin_fichier_stat)
        layout_stats.addWidget(self.Confirm_select_fichier_stat)
        layout_stats.addWidget(self.label_select_col_ref)
        layout_stats.addWidget(self.Cb_feuille_ref)
        layout_stats.addWidget(self.Cb_Colonne_ref)
        layout_stats.addWidget(self.label_select_colonnes_secondaires)
        layout_stats.addWidget(self.Cb_feuille_coll_secondaire)
        layout_stats.addWidget(self.Cb_coll_secondaire)
        layout_stats.addWidget(self.Button_ajouter_coll_secondaire)
        layout_stats.addWidget(self.Button_supprimer_coll_secondaire)
        layout_stats.addWidget(Scroll_Area_coll_secondaire)
        layout_stats.addWidget(self.coched_exclure_datas)
        layout_stats.addWidget(self.label_nom_feuille_created)
        layout_stats.addWidget(self.nom_feuille_created)
        layout_stats.addWidget(self.Button_appliquer)

#------------------------3EME ONGLET--> insertion de colonne d'un fichier à un autre en fonction des numéros de trades---------
        self.select_first_fichier = QPushButton("Selectionner Fichier")
        self.select_first_fichier.clicked.connect(self.showFileDialog)
        self.Confirm_select_fichier_first = QPushButton("Confirme la selection du fichier")
        self.Confirm_select_fichier_first.clicked.connect(self.extraction_critères)
        self.chemin_fichier_first = QLineEdit()
        self.chemin_fichier_first.setReadOnly(True)
        self.label_select_col_fichier_first = QLabel("Selection de colonnes à transferer vers la feuille d'un autre fichier:")
        self.Cb_feuille_fichier_first = QComboBox()
        self.Cb_Colonne_fichier_first = QComboBox()
        self.Cb_feuille_fichier_first.currentTextChanged.connect(self.change_combobox_Feuille)
        self.Button_ajouter_coll_fichier_first = QPushButton("+Ajouter colonne à insérer+")
        self.Button_ajouter_coll_fichier_first.clicked.connect(self.Ajouter_colonne_fichier_first)
        self.Button_supprimer_fichier_first = QPushButton("-Supprimer colonne secondaire-")
        self.Button_supprimer_fichier_first.clicked.connect(self.remove_last_widget_scroll_area_fichier_first)

        # Creation de la Scroll Area
        Scroll_Area_fichier_first = QScrollArea(self)
        widget_container_fichier_first = QWidget(self)
        Scroll_Area_fichier_first.setWidget(widget_container_fichier_first)
        Scroll_Area_fichier_first.setWidgetResizable(True)
        self.scroll_layout_fichier_first = QGridLayout(widget_container_fichier_first)

        self.second_fichier_select=QPushButton("Selectionner Fichier pour écrire les colonnes choisient")
        self.second_fichier_select.clicked.connect(self.showFileDialog)
        self.Confirmer_second_fichier_select=QPushButton("Confirmation selection fichier")
        self.Confirmer_second_fichier_select.clicked.connect(self.extraction_critères)
        self.chemin_fichier_second = QLineEdit()
        self.chemin_fichier_second.setReadOnly(True)
        self.label_feuille_fichier_second=QLabel("Selection de la feuille où appliquer les colonnes choisient:")
        self.feuille_fichier_second=QComboBox()
        self.Button_appliquer_merge_coll = QPushButton("Appliquer")
        self.Button_appliquer_merge_coll.setStyleSheet("""background-color:red""")
        self.Button_appliquer_merge_coll.clicked.connect(self.Traitement_onglet_3)

        layout_transfert_F_to_F.addWidget(self.select_first_fichier)
        layout_transfert_F_to_F.addWidget(self.Confirm_select_fichier_first)
        layout_transfert_F_to_F.addWidget(self.chemin_fichier_first)
        layout_transfert_F_to_F.addWidget(self.label_select_col_fichier_first)
        layout_transfert_F_to_F.addWidget(self.Cb_feuille_fichier_first)
        layout_transfert_F_to_F.addWidget(self.Cb_Colonne_fichier_first)
        layout_transfert_F_to_F.addWidget(self.Button_ajouter_coll_fichier_first)
        layout_transfert_F_to_F.addWidget(self.Button_supprimer_fichier_first)
        layout_transfert_F_to_F.addWidget(Scroll_Area_fichier_first)
        layout_transfert_F_to_F.addWidget(self.second_fichier_select)
        layout_transfert_F_to_F.addWidget(self.Confirmer_second_fichier_select)
        layout_transfert_F_to_F.addWidget(self.chemin_fichier_second)
        layout_transfert_F_to_F.addWidget(self.label_feuille_fichier_second)
        layout_transfert_F_to_F.addWidget(self.feuille_fichier_second)
        layout_transfert_F_to_F.addWidget(self.Button_appliquer_merge_coll)

# ------------------Application du style au widgets--------------------
        #Style feuille
        self.setStyleSheet(style_beige)
        #Style onglet
        self.tabs.setStyleSheet(style_onglet)

        #Style widget Button + label
        liste_btn_supprimer=[self.Btn_Supprimer_feuille,self.Btn_data_remove_extraire,self.Btn_remove_last_critère,
                             self.Button_supprimer_coll_secondaire,self.Button_supprimer_fichier_first]
        liste_btn_ajouter=[self.Btn_ajouter_critères,self.Btn_Ajouter_feuille,self.Button_ajouter_coll_secondaire,
                           self.Button_ajouter_coll_fichier_first,self.Btn_data_extraire]
        liste_btn_confirmer=[self.Confirmer_second_fichier_select,self.Confirm_select_fichier_first,
                             self.Confirm_select_fichier_stat,self.select_fichier_stat,self.select_fichier_stat,
                             self.select_first_fichier,self.select_file_button,self.selection_result_folder,
                             self.Button_extraction_feuilles_critères,self.button_select_folder,self.second_fichier_select]
        liste_btn_appliquer=[self.Button_appliquer,self.Button_appliquer_merge_coll,
                             self.button_appliquer_to_excel]
        liste_label=[self.label_select_colonnes_secondaires,self.label_select_col_ref,self.label_select_col_fichier_first,
                     self.label_feuille_fichier_second,self.label_feuilles_created,self.label_name_sheet_excel,self.Label_extraire,
                     self.Label_ajout_critère,self.file_path_label,self.label_name_fichier_excel,self.label_nom_feuille_created]
        liste_check_box=[self.coched_exclure_datas]
        liste_radio_box=[self.radio_stats_ref,self.radio_stats_without_ref]
        for btn in liste_btn_supprimer:
            btn.setStyleSheet(style_btn_supprimer)
        for btn in liste_btn_ajouter:
            btn.setStyleSheet(style_btn_ajouter)
        for btn in liste_btn_confirmer:
            btn.setStyleSheet(style_btn_selection)
        for btn in liste_btn_appliquer:
            btn.setStyleSheet(style_btn_traitement)
        for label in liste_label:
            label.setStyleSheet(style_label)
        for checkbox in liste_check_box:
            checkbox.setStyleSheet(style_checkbox)
        for radiobox in liste_radio_box:
            radiobox.setStyleSheet(style_radio_box)








    def Ajouter_feuille(self):
        dicos_c=extraire_label_dico_scroll_area(self.scroll_layout)
        dico_critère=dicos_c["liste_dico"]
        dicos_datas_extraire=extraire_label_dico_scroll_area(self.scroll_layout_extraire)
        dico_données_extraire=dicos_datas_extraire["liste_dico"]
        nom_feuille=self.Name_sheet.text()
        dico_feuille=f"{{'feuille':'{nom_feuille}','critères':{dico_critère},'datas_extraire':{dico_données_extraire}}}"
        self.scroll_layout_feuilles_created.addWidget(QLabel(dico_feuille))

    def Supprimer_feuille(self):
        if self.scroll_layout_feuilles_created.count() > 0:
            item = self.scroll_layout_feuilles_created.itemAt(self.scroll_layout_feuilles_created.count() - 1)
            if item:
                widget = item.widget()
                if widget:
                    widget.deleteLater()
    def change_widget_radio_stats(self):
        if self.sender()==self.radio_stats_ref:
            self.Button_supprimer_coll_secondaire.setText("-Supprimer colonne secondaire-")
            self.Button_ajouter_coll_secondaire.setText("+Ajouter colonne secondaire+")
            self.label_select_colonnes_secondaires.setText("Selection des colonnes secondaires:")
            self.label_select_col_ref.setVisible(True)
            self.Cb_Colonne_ref.setVisible(True)
            self.Cb_feuille_ref.setVisible(True)
        elif self.sender()==self.radio_stats_without_ref:
            self.Button_ajouter_coll_secondaire.setText("+Ajouter colonne+")
            self.Button_supprimer_coll_secondaire.setText("-Supprimer colonne-")
            self.label_select_colonnes_secondaires.setText("Selection des colonnes")
            self.Cb_Colonne_ref.setVisible(False)
            self.label_select_col_ref.setVisible(False)
            self.Cb_feuille_ref.setVisible(False)



    def Traitement_onglet_3(self):
        liste_colonnes=scroll_area_json_to_list_dico(self.scroll_layout_fichier_first)
        chemin_fichier_colonne=self.chemin_fichier_first.text()
        chemin_fichier_second=self.chemin_fichier_second.text()
        feuille_fichier_second=self.feuille_fichier_second.currentText()
        df_coll=extration_colonnes(chemin_fichier_colonne,liste_colonnes)
        df_main=pd.read_excel(chemin_fichier_second,sheet_name=feuille_fichier_second,engine="openpyxl")
        df_trade=find_colonne_trade(df_main)
        nouveau_df = pd.merge(df_coll,df_trade,on="Trade",how='inner')
        ecriture_df_to_excel_col_suivante(chemin_fichier_second,feuille_fichier_second,nouveau_df)



    def Ajouter_colonne_fichier_first(self):
        Feuille = self.Cb_feuille_fichier_first.currentText()
        Colonne = self.Cb_Colonne_fichier_first.currentText()

        Json_condition = f"{{\"Feuille\":\"{Feuille}\",\"Colonne\":\"{Colonne}\"}}"


        Condition = QLabel(Json_condition)
        self.scroll_layout_fichier_first.addWidget(Condition)

    def remove_last_widget_scroll_area_fichier_first(self):
        if self.scroll_layout_fichier_first.count() > 0:
            item = self.scroll_layout_fichier_first.itemAt(self.scroll_layout_fichier_first.count() - 1)
            if item:
                widget = item.widget()
                if widget:
                    widget.deleteLater()

    def Construire_appliquer_stats(self):
        if self.radio_stats_ref.isChecked():
            chemin=self.chemin_fichier_stat.text()
            feuille_ref=self.Cb_feuille_ref.currentText()
            col_ref=self.Cb_Colonne_ref.currentText()
            feuille_sec=self.Cb_feuille_coll_secondaire.currentText()
            if feuille_ref!=feuille_sec:
                print("ERREUR")
            liste_label=liste_label=extraire_label_dico_scroll_area(self.scroll_layout_coll_secondaire)
            dico_possibility_col_ref=extraction_colonne_ref_possibilité(chemin,feuille_ref,col_ref)
            dico_possibility_coll_sec=create_dico_différentes_coll(liste_label["liste_widget"],chemin)
            all_combinaisons_coll_sec=All_combinaisons_matrice_dictionnary(dico_possibility_coll_sec)
            df=compter_possibilité_coll_ref(chemin,feuille_ref,all_combinaisons_coll_sec,dico_possibility_col_ref)
            name_feuille=self.nom_feuille_created.text()
            création_sheet_excel(chemin=chemin,df=df,name_feuille=name_feuille)
        elif self.radio_stats_without_ref.isChecked():
            chemin=self.chemin_fichier_stat.text()
            exclure_data_vide_and_more=self.coched_exclure_datas.isChecked()
            liste_label=extraire_label_dico_scroll_area(self.scroll_layout_coll_secondaire)
            name_feuille=self.nom_feuille_created.text()
            df=fusion_clean_dicos(chemin=chemin,liste_dico_feuille=liste_label["liste_dico"],Boolean_exlure_datas=exclure_data_vide_and_more)
            # Supprimez la colonne "Trade" du DataFrame pour extraire les possibilités plus facilement
            df.drop(columns=['Trade'], inplace=True)
            dico_possibilities_colonnes=extration_all_possibilities_without_coll_trade(df)
            all_possibilies=All_combinaisons_matrice_dictionnary(dico_possibilities_colonnes)
            df_stat=stats_without_col_ref(all_possibilies,df)
            création_sheet_excel(chemin,df_stat,name_feuille)


            #Ok donc j'ai clean les datas du df, il me reste à extraire les possibilités du DF pour chaque colonne à part Trade
            #création_sheet_excel(chemin=chemin,df=df,name_feuille=name_feuille)





    #Fonction pour afficher et changer value default excel
    def fct_change_value_trie_filtrage(self):
        self.w = Change_value_filtre_tri_excel_window()
        self.resize(300,300)
        self.w.show()



    def remove_last_widget_scroll_area(self):
        if self.scroll_layout.count() > 0:
            item = self.scroll_layout.itemAt(self.scroll_layout.count() - 1)
            if item:
                widget = item.widget()
                if widget:
                    widget.deleteLater()
    def remove_last_widget_scroll_area_extraire(self):
        if self.scroll_layout_extraire.count() > 0:
            item = self.scroll_layout_extraire.itemAt(self.scroll_layout_extraire.count() - 1)
            if item:
                widget = item.widget()
                if widget:
                    widget.deleteLater()

    def remove_last_widget_scroll_area_coll_secondaires(self):
        if self.scroll_layout_coll_secondaire.count() > 0:
            item = self.scroll_layout_coll_secondaire.itemAt(self.scroll_layout_coll_secondaire.count() - 1)
            if item:
                widget = item.widget()
                if widget:
                    widget.deleteLater()


    def Ajouter_critères(self):
        Feuille=self.Cb_feuille_critère.currentText()
        Colonne=self.Cb_critères.currentText()
        opérateur=self.Opérateurs.currentText()
        Valeur=self.Comparaison.text()
        if Valeur.isdigit():
            Json_condition=f"{{\"Feuille\":\"{Feuille}\",\"Colonne\":\"{Colonne}\",\"Opérateur\":\"{opérateur}\",\"Valeur\":{Valeur}}}"
        else:
            Json_condition=f"{{\"Feuille\":\"{Feuille}\",\"Colonne\":\"{Colonne}\",\"Opérateur\":\"{opérateur}\",\"Valeur\":\"{Valeur}\"}}"

        Condition=QLabel(Json_condition)
        self.scroll_layout.addWidget(Condition)

    def Ajouter_data_extraire(self):

        feuille=self.Cb_feuille_extraire.currentText()
        colonne=self.Cb_colonne_extraire.currentText()
        text_label=f"{{\"Feuille\":\"{feuille}\",\"Colonne\":\"{colonne}\"}}"
        data_append=QLabel(text_label)
        self.scroll_layout_extraire.addWidget(data_append)

    def Ajouter_colonne_secondaires(self):
        feuille=self.Cb_feuille_coll_secondaire.currentText()
        colonne=self.Cb_coll_secondaire.currentText()
        text_label = f"{{\"Feuille\":\"{feuille}\",\"Colonne\":\"{colonne}\"}}"
        data_append = QLabel(text_label)
        self.scroll_layout_coll_secondaire.addWidget(data_append)


    #Permet d'extracter tout les entetes et les pages du fichier
    #Puis les appliques dans les différentes QCombobox
    def extraction_critères(self):
        if self.sender()==self.Button_extraction_feuilles_critères:
            chemin=self.file_path_label.text()
        elif self.sender()==self.Confirm_select_fichier_stat:
            chemin=self.chemin_fichier_stat.text()
        elif self.sender()==self.Confirm_select_fichier_first:
            chemin=self.chemin_fichier_first.text()
        elif self.sender()==self.Confirmer_second_fichier_select:
            chemin=self.chemin_fichier_second.text()
        # Chargez le classeur Excel
        workbook = openpyxl.load_workbook(chemin)

        # Récupérez la liste des noms des feuilles
        sheet_names = workbook.sheetnames

        # Parcourez toutes les feuilles et obtenez les en-têtes
        dico_sheet={}
        liste_sheet_name=[]
        for sheet_name in sheet_names:
            sheet = workbook[sheet_name]
            headers = [cell.value for cell in sheet[1]]
            dico_sheet[sheet_name]=headers
            liste_sheet_name.append(sheet_name)
        if self.sender()==self.Button_extraction_feuilles_critères:
            self.Cb_feuille_critère.clear()
            self.Cb_feuille_critère.addItems(liste_sheet_name)
            self.Cb_feuille_extraire.clear()
            self.Cb_feuille_extraire.addItems(liste_sheet_name)
            # Charger le contenu actuel du fichier JSON
            with open("C:\\Users\\Baptiste\\Documents\\GUI_ext_tri_datas\\données.json", "r") as json_file:
                data = json.load(json_file)

            # Mettre à jour le dictionnaire "trie_filtre_temp" avec le contenu de dico_sheet
            data["trie_filtre_temp"] = dico_sheet

            # Sauvegarder le fichier JSON mis à jour
            with open("C:\\Users\\Baptiste\\Documents\\GUI_ext_tri_datas\\données.json", "w") as json_file:
                json.dump(data, json_file, indent=4)
        elif self.sender()==self.Confirm_select_fichier_stat:
            self.Cb_feuille_coll_secondaire.clear()
            self.Cb_feuille_ref.clear()
            self.Cb_feuille_coll_secondaire.addItems(liste_sheet_name)
            self.Cb_feuille_ref.addItems(liste_sheet_name)

            with open("C:\\Users\\Baptiste\\Documents\\GUI_ext_tri_datas\\données.json", "r") as json_file:
                data = json.load(json_file)
            data["stats_temp"] = dico_sheet
            with open("données.json", "w") as json_file:
                json.dump(data, json_file, indent=4)
        elif self.sender()==self.Confirm_select_fichier_first:
            self.Cb_feuille_fichier_first.clear()
            self.Cb_Colonne_fichier_first.clear()
            self.Cb_feuille_fichier_first.addItems(liste_sheet_name)


            with open("C:\\Users\\Baptiste\\Documents\\GUI_ext_tri_datas\\données.json","r") as json_file:
                data=json.load(json_file)
            data["F_to_F_temp"]=dico_sheet
            with open("données.json", "w") as json_file:
                json.dump(data, json_file, indent=4)

        elif self.sender()==self.Confirmer_second_fichier_select:
            self.feuille_fichier_second.clear()
            self.feuille_fichier_second.addItems(liste_sheet_name)




    def change_combobox_Feuille(self):

        page=self.sender().currentText()
        if self.sender()==self.Cb_feuille_extraire or self.sender()==self.Cb_feuille_critère:
            with open("C:\\Users\\Baptiste\\Documents\\GUI_ext_tri_datas\\données.json", "r") as json_file:
                data = json.load(json_file)
            feuille_data = data["trie_filtre_temp"][page]
            if self.sender()==self.Cb_feuille_critère:
                self.Cb_critères.clear()
                self.Cb_critères.addItems(feuille_data)
            elif self.sender()==self.Cb_feuille_extraire:
                self.Cb_colonne_extraire.clear()
                self.Cb_colonne_extraire.addItems(feuille_data)
        elif self.sender()==self.Cb_feuille_ref or self.sender()==self.Cb_feuille_coll_secondaire:
            with open("C:\\Users\\Baptiste\\Documents\\GUI_ext_tri_datas\\données.json", "r") as json_file:
                data = json.load(json_file)
            feuille_data = data["stats_temp"][page]
            if self.sender()==self.Cb_feuille_ref:
                self.Cb_Colonne_ref.clear()
                self.Cb_Colonne_ref.addItems(feuille_data)
            elif self.sender()==self.Cb_feuille_coll_secondaire:
                self.Cb_coll_secondaire.clear()
                self.Cb_coll_secondaire.addItems(feuille_data)
        elif self.sender()==self.Cb_feuille_fichier_first:
            with open("C:\\Users\\Baptiste\\Documents\\GUI_ext_tri_datas\\données.json","r") as json_file:
                data=json.load(json_file)
            feuille_data =data["F_to_F_temp"][page]
            self.Cb_Colonne_fichier_first.clear()
            self.Cb_Colonne_fichier_first.addItems(feuille_data)

    # Pour selectionner un dossier
    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Sélectionner un Dossier", "/path/par/défaut")
        self.selection_result_folder.setText(folder)
    #Pour selectionner un fichier
    def showFileDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly

        file_dialog = QFileDialog(self)
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        file_dialog.setOptions(options)

        file_path, _ = file_dialog.getOpenFileName(self, "Sélectionner un fichier", "", "Tous les fichiers (*)")
        if self.sender()==self.select_file_button:
            if file_path:
                self.file_path_label.setText(file_path)
            else:
                self.file_path_label.setText("")
        elif self.sender()==self.select_fichier_stat:
            if file_path:
                self.chemin_fichier_stat.setText(file_path)
            else:
                self.chemin_fichier_stat.setText("")
        elif self.sender()==self.select_first_fichier:
            if file_path:
                self.chemin_fichier_first.setText(file_path)
            else:
                self.chemin_fichier_first.setText("")
        elif self.sender()==self.second_fichier_select:
            if file_path:
                self.chemin_fichier_second.setText(file_path)
            else:
                self.chemin_fichier_second.setText("")


    def Traitement_données(self):
        chemin_dest_final = f"{self.selection_result_folder.text()}/{self.name_fichier_excel.text()}.xlsx"
        dicos_data=extraire_label_dico_scroll_area_special(self.scroll_layout_feuilles_created)
        dicos_feuilles=dicos_data["liste_dico"]
        for dico_feuille in dicos_feuilles:
            #dico_feuille={"feuille":nom_feuille,"critères":dico_critère,"datas_extraire":dico_données_extraire}
            name_sheet=dico_feuille['feuille']
            liste_critères=dico_feuille["critères"]
            liste_datas_extraire=dico_feuille["datas_extraire"]
            #Traitement data et construction du excel
            chemin=self.file_path_label.text()
            serie_nb_trade=traitement_critères(liste_critères,chemin)
            df_final,colonne_screenshots=extraire_datas(chemin,liste_datas_extraire,serie_nb_trade)

            if not os.path.exists(chemin_dest_final):
                # Si le fichier n'existe pas, créez-le en sauvegardant le DataFrame complet
                df_final.to_excel(chemin_dest_final, index=False, engine="openpyxl", sheet_name=name_sheet)
            else:
                # Ouvrez le fichier Excel en mode append
                with pd.ExcelWriter(chemin_dest_final, mode='a', engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False, sheet_name=name_sheet)

            #df_final.to_excel(chemin_dest_final,index=False,engine="openpyxl",sheet_name=name_sheet)  # Utilisez index=False si vous ne souhaitez pas enregistrer l'index du DataFrame
            écrire_colonne_hyperlink(chemin,colonne_screenshots,chemin_dest_final,serie_nb_trade,name_sheet) #écrires les colonnes de screen en hyperlink









def main():
    app = QApplication(sys.argv)
    window = FilePickerApp()
    window.showMaximized()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()