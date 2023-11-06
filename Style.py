
# Style pour les Boutons des pages principales
style_Selection = """QPushButton{background-color:#3F57F6;color:#1F2F6F;font-weight: bold;font-size: 16px;border: 2px groove gray}"""
style_Selection_pressed = """QPushButton:pressed{background-color:#2D44C5;color:#1F2F6F;font-weight: bold;font-size: 16px;border: 2px groove gray}"""
style_Selection_hover = """QPushButton:hover{background-color: #5C6BF6;color: #1F2F6F;font-weight: bold;font-size: 16px;border: 2px groove gray}"""
style_btn_selection = style_Selection + style_Selection_hover+style_Selection_pressed

style_ajouter = """ QPushButton{background-color: #F2ABCE; border-radius: 30px; color: #2698A5;font-weight: bold;}"""
style_ajouter_pressed = """QPushButton:pressed{background-color: #D58FAF; border-radius: 30px; color: #2698A5;font-weight: bold;}"""
style_ajouter_hover = """QPushButton:hover{ background-color: #FFC0CB;border-radius: 30px;color: #2698A5;font-weight: bold;}"""
style_btn_ajouter = style_ajouter + style_ajouter_hover+style_ajouter_pressed


style_supprimer = """ QPushButton{background-color:#DFD547; border-radius: 30px; color: #536427;font-weight: bold;}"""
style_supprimer_pressed = """QPushButton:pressed{background-color:#BFB83A; border-radius: 30px; color: #536427;font-weight: bold;}"""
style_supprimer_hover = """QPushButton:hover{background-color:#EDE380; border-radius: 30px; color: #536427;font-weight: bold;}"""
style_btn_supprimer = style_supprimer+style_supprimer_hover+style_supprimer_pressed

style_traitement="""QPushButton{background-color:#D70A47;color:#1F2F6F;font-weight: bold;font-size: 20px;border: 2px groove gray}"""
style_traitement_hover="""QPushButton:hover{background-color:#FF4A73 ;color: #1F2F6F;font-weight: bold;font-size: 20px;border: 2px groove gray}"""
style_traitement_pressed="""QPushButton:pressed{background-color: #AD062E;color: #1F2F6F;font-weight: bold;font-size: 20px;border: 2px groove gray}"""
style_btn_traitement=style_traitement+style_traitement_hover+style_traitement_pressed

style_label="""QLabel{
    font-family: Arial;
    font-size: 18px;
    color: #3498db;
    border-radius: 5px;
    font-weight: bold;
}"""


style_checkbox_normal="""QCheckBox { color: black;font-weight:bold}"""
style_checkbox_indicator="""QCheckBox::indicator { background-color: white; border: 2px solid #D70A47; border-radius: 3px; }"""
style_checkbox_indicator_checked="""QCheckBox::indicator:checked { background-color:#D70A47; border: 2px solid #D70A47; border-radius: 3px; }"""
style_checkbox=style_checkbox_normal+style_checkbox_indicator+style_checkbox_indicator_checked

style_radio_box_normal="""QRadioButton { background: none; border: none; color: #333;font-weight:bold }"""
style_radio_box_indicator="""QRadioButton::indicator {background-color: white; border: 2px solid #D70A47; border-radius: 3px; }"""
style_radio_box_indicator_checked="""QRadioButton::indicator:checked { background-color:#D70A47; border: 2px solid #D70A47; border-radius: 3px; }"""
style_radio_box=style_radio_box_normal+style_radio_box_indicator+style_radio_box_indicator_checked

#Style feuille
style_beige="background-color: #faf3e0;"

style_onglet_normal="""QTabBar::tab { background-color: beige;color: #333; font-size: 16px;min-height: 200px;  } """
style_onglet_selected="""QTabBar::tab:selected { background-color: lightblue; font-weight: bold;min-height: 200px;  }"""
style_onglet=style_onglet_normal+style_onglet_selected
