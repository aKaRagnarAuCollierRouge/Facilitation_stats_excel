import json
import os

from PySide2.QtWidgets import QWidget, QGridLayout, QScrollArea, QPushButton, QLineEdit, QComboBox, QLabel


class Change_value_filtre_tri_excel_window(QWidget):

    def __init__(self):
        super().__init__()
        #Telecharger datas:
        # Emplacement du répertoire du script Python
        emplacement_script = os.path.dirname(os.path.abspath(__file__))
        # Nom de votre fichier JSON
        nom_fichier_json = "données.json"
         # Construction du chemin relatif vers le répertoire parent
        self.chemin_json = os.path.join(emplacement_script, "..", nom_fichier_json)

        with open(self.chemin_json, "r") as fichier_json:
            data = json.load(fichier_json)
        screen_words = data["valeurs_defaults"]["Screen_words"]
        col_ref=data["valeurs_defaults"]["col_ref_words"]

        #Mise en place des widgets
        self.setWindowTitle("Changer value default filtrage-trie")
        layout = QGridLayout()
        self.label_hyperlink=QLabel("Ajouter mot référencer colonne hyperlink")
        self.word_hyperlink=QLineEdit()
        self.Btn_ajouter_world_hyperlink=QPushButton("Ajouter mot Hyperlien")
        self.Btn_ajouter_world_hyperlink.clicked.connect(self.Ajouter_mot)
        self.Cb_liste_world_hyperlink_supprimer=QComboBox()
        self.Cb_liste_world_hyperlink_supprimer.addItems(screen_words)
        self.Btn_supprimer_world_hyperlink=QPushButton('Supprimer mot hyperlink')
        self.Btn_supprimer_world_hyperlink.clicked.connect(self.Supprimer_mot)

        self.label_colonne_Trade=QLabel("Ajouter|Supprimer mots qui peut désigner colonne de référence")
        self.Btn_ajouter_word_coll_ref=QPushButton("Ajouter Mot")
        self.Btn_ajouter_word_coll_ref.clicked.connect(self.Ajouter_mot)
        self.word_coll_ref = QLineEdit()
        self.Cb_supprimer_word_coll_ref=QComboBox()
        self.Cb_supprimer_word_coll_ref.addItems(col_ref)
        self.Btn_supprimer_word_ref=QPushButton("Supprimer Mot")
        self.Btn_supprimer_word_ref.clicked.connect(self.Supprimer_mot)


        layout.addWidget(self.label_hyperlink)
        layout.addWidget(self.word_hyperlink)
        layout.addWidget(self.Btn_ajouter_world_hyperlink)
        layout.addWidget(self.Cb_liste_world_hyperlink_supprimer)
        layout.addWidget(self.Btn_supprimer_world_hyperlink)

        layout.addWidget(self.label_colonne_Trade)
        layout.addWidget(self.Btn_ajouter_word_coll_ref)
        layout.addWidget(self.word_coll_ref)
        layout.addWidget(self.Cb_supprimer_word_coll_ref)
        layout.addWidget(self.Btn_supprimer_word_ref)
        self.setLayout(layout)

    def Ajouter_mot(self):
        with open(self.chemin_json, "r") as fichier_json:
            data = json.load(fichier_json)

        if self.sender()==self.Btn_ajouter_world_hyperlink:
            Mot_ajouté=self.word_hyperlink.text()
            data["valeurs_defaults"]["Screen_words"].append(Mot_ajouté)
            self.Cb_liste_world_hyperlink_supprimer.clear()
            self.Cb_liste_world_hyperlink_supprimer.addItems(data["valeurs_defaults"]["Screen_words"])
        elif self.sender()==self.Btn_ajouter_word_coll_ref:
            Mot_ajouté=self.word_coll_ref.text()
            data["valeurs_defaults"]["col_ref_words"].append(Mot_ajouté)
            self.Cb_supprimer_word_coll_ref.clear()
            self.Cb_supprimer_word_coll_ref.addItems(data["valeurs_defaults"]["col_ref_words"])

        # Sauvegarder les modifications dans le fichier JSON
        with open(self.chemin_json, "w") as fichier_json:
            json.dump(data, fichier_json, indent=4)

    def Supprimer_mot(self):
        with open(self.chemin_json, "r") as fichier_json:
            data = json.load(fichier_json)

        if self.sender() == self.Btn_supprimer_world_hyperlink:
            Mot_supprimé = self.Cb_liste_world_hyperlink_supprimer.currentText()
            data["valeurs_defaults"]["Screen_words"].remove(Mot_supprimé)
            self.Cb_liste_world_hyperlink_supprimer.clear()
            self.Cb_liste_world_hyperlink_supprimer.addItems(data["valeurs_defaults"]["Screen_words"])
        elif self.sender() == self.Btn_supprimer_word_ref:
            Mot_supprimé = self.Cb_supprimer_word_coll_ref.currentText()
            data["valeurs_defaults"]["col_ref_words"].remove(Mot_supprimé)
            self.Cb_supprimer_word_coll_ref.clear()
            self.Cb_supprimer_word_coll_ref.addItems(data["valeurs_defaults"]["col_ref_words"])
        # Sauvegarder les modifications dans le fichier JSON
        with open(self.chemin_json, "w") as fichier_json:
            json.dump(data, fichier_json, indent=4)






