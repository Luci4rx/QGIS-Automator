#! python3  # noqa: E265

"""
    Main plugin module.
"""


# __Стандартні модулі середовища пайтон__
import re
from functools import partial
from pathlib import Path

# __Модулі QGIS__
from qgis.core import QgsApplication, QgsSettings, QgsProject, QgsFeatureRequest, QgsRectangle, QgsLayoutSize,  QgsUnitTypes, QgsMapLayer, QgsExpression # type: ignore
from qgis.gui import QgisInterface # type: ignore
from qgis.PyQt.QtCore import QCoreApplication, QLocale, QTranslator, QUrl, Qt, NULL, QDate # type: ignore
from qgis.PyQt.QtGui import QDesktopServices, QIcon # type: ignore
from qgis.PyQt.QtWidgets import QAction, QLineEdit, QDialog, QVBoxLayout, QPushButton, QTextEdit, QLabel, QDockWidget, QWidget, QTableWidget, QTableWidgetItem, QComboBox, QMainWindow, QHBoxLayout,QCheckBox # type: ignore

#  __Підключення модуля ресурсів__
from .resources import *

# __Додаткові бібліотеки для роботи плагіна__
import pymorphy2 # type: ignore

# __Стандартні бібліотеки шаблонізатора плагінів__
from ops_automator.__about__ import ( # type: ignore
    DIR_PLUGIN_ROOT,
    __icon_path__,
    __title__,
    __uri_homepage__,
) 
from ops_automator.gui.dlg_settings import PlgOptionsFactory # type: ignore
from ops_automator.toolbelt import PlgLogger # type: ignore

##################################
########## Classes ###############
##################################

# -- Клас діалогового вікна -- #
class MainDialog(QDialog):
    def __init__(self, text):
        super().__init__()

        self.setWindowTitle("Ops Automator")
        layout = QVBoxLayout(self)
        button = QPushButton("Ок")
        label = QLabel(text)
        button.clicked.connect(self.on_button_clicked)
        layout.addWidget(label)
        layout.addWidget(button)

    # Закриття діалогового вікна при натисканні кнопки
    def on_button_clicked(self):
        self.close()  

# -- Клас таблиці пошуку по ПІБ якщо більше одного власника -- #
class InfoSerchTable(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Інформація про об'єкти")
        layout = QVBoxLayout(self)
        self.table_widget = QTableWidget()
        self.table_widget.setColumnCount(2)
        self.table_widget.setHorizontalHeaderLabels(["Назва власника", "Кадастровий номер"])
        layout.addWidget(self.table_widget)
        self.setLayout(layout)

# -- Клас таблиці атрибутів вибраного об'єкта -- #
class InfoTablePanel(QWidget):
    def __init__(self, iface):
        super().__init__()
        self.iface = iface
        self.initUI()

    # Опис інтерфейсу віджета таблиці атрибутів
    def initUI(self):
        layout = QVBoxLayout()
        self.reftable = QPushButton("Оновити")
        self.reftable.clicked.connect(self.RefreshTableInfo)
        layout.addWidget(self.reftable)
        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Атрибут", "Значення"])
        layout.addWidget(self.table)
        self.setLayout(layout)

    # Оновлення таблиці атрибутів
    def RefreshTableInfo(self):
        self.table.setRowCount(0)
        layer = self.iface.activeLayer()
        if layer is not None:
            selected_features = layer.selectedFeatures()
            if selected_features:
                for feature in selected_features:
                    for field in layer.fields():
                        field_value = feature[field.name()]
                        if field_value != NULL:
                                field_alias = field.alias()
                                if field_alias == '':
                                    field_alias= field.name()
                                rowPosition = self.table.rowCount()
                                self.table.insertRow(rowPosition)
                                for i, item in enumerate([field_alias , field_value]):
                                    if isinstance(item, QtCore.QDate):
                                        self.table.setItem(rowPosition, i, QTableWidgetItem(item.toString("yyyy-MM-dd")))
                                    else:
                                        self.table.setItem(rowPosition, i, QTableWidgetItem(str(item)))
            else:
                    dialog = MainDialog("Ділянка не вибрана")
                    dialog.exec_()
        else:
            print("Шар не обраний")

# -- Клас вікна налаштування фільтрування -- #
class FilterDataManage(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Фільтр")
        self.setGeometry(100, 100, 800, 600)
        self.initUI()
    def initUI(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout()  
        central_widget.setLayout(main_layout)
        combo_filter_layout = QHBoxLayout()
        main_layout.addLayout(combo_filter_layout)
        self.combo_box1 = QComboBox()
        self.combo_box1.setMinimumWidth(300)
        combo_filter_layout.addWidget(self.combo_box1)
        self.populateVectorLayersCombo()
        self.combo_box2 = QComboBox()
        self.combo_box2.setMinimumWidth(300)
        combo_filter_layout.addWidget(self.combo_box2)
        self.combo_box1.currentIndexChanged.connect(self.populateFieldNamesCombo)
        self.combo_box1.currentIndexChanged.connect(self.showAttributes)
        self.line_edit = QLineEdit()
        self.line_edit.setPlaceholderText("Введіть значення фільтра")
        combo_filter_layout.addWidget(self.line_edit)
        self.filter_button = QPushButton("Фільтрувати")
        self.filter_button.clicked.connect(self.ParseInput)
        combo_filter_layout.addWidget(self.filter_button)
        table_layout = QHBoxLayout()
        main_layout.addLayout(table_layout)
        self.table = QTableWidget()
        table_layout.addWidget(self.table)
        self.table_orign = QTableWidget()
        table_layout.addWidget(self.table_orign)

    def populateVectorLayersCombo(self):
        self.combo_box1.clear()
        layers = QgsProject.instance().mapLayers().values()
        for layer in layers:
            if layer.type() == QgsMapLayer.VectorLayer:
                self.combo_box1.addItem(layer.name())

    def populateFieldNamesCombo(self):
        layer_name = self.combo_box1.currentText()
        layer = QgsProject.instance().mapLayersByName(layer_name)[0]
        self.combo_box2.clear()
        fields = layer.fields()
        for field in fields:

            if field.alias() == '':
                self.combo_box2.addItem(field.name())
            else:
                self.combo_box2.addItem(field.alias())

    def FilterByExpression(self, expression_text, layer_name):
        layer = QgsProject.instance().mapLayersByName(layer_name)[0]
        expression = QgsExpression(expression_text)
        if not expression.hasParserError():
            request = QgsFeatureRequest().setFilterExpression(expression_text)
            self.FillFiltrTable(layer, request)
        else:
            print("Помилка виразу:", expression.parserErrorString())

    def FillFiltrTable(self, layer, req):
        self.table_orign.clear()
        features = layer.getFeatures(req)
        fields = layer.fields()
        self.table_orign.setRowCount(len(list(layer.getFeatures(req))))
        self.table_orign.setColumnCount(len(fields))
        for i, field in enumerate(fields):
            
            if field.alias() == "":
                self.table_orign.setHorizontalHeaderItem(i, QTableWidgetItem(field.name()))
            else:
                self.table_orign.setHorizontalHeaderItem(i, QTableWidgetItem(field.alias()))
        for i, feature in enumerate(features):
            for j, field in enumerate(fields):
                if isinstance(feature[field.name()], QtCore.QDate):
                    self.table_orign.setItem(i, j, QTableWidgetItem(feature[field.name()].toString("yyyy-MM-dd")))
                else:
                    self.table_orign.setItem(i, j, QTableWidgetItem(str(feature[field.name()])))
                
    def ParseInput(self):
        filter_value = self.line_edit.text()
        layer_name = self.combo_box1.currentText()
        filed = self.combo_box2.currentText()

        def GetNemIntFloatText():
            getnum_reg = r"-?\d*\.?\d+"
            num_list = []
            text = re.findall(getnum_reg, filter_value)
            for i in text:
                num_list.append(float(i))
            return num_list
        if re.fullmatch(r"<\s*-?\d+(\.\d+)?\s*", filter_value):
            res = GetNemIntFloatText()
            expression_text = f""" "{filed}" < {res[0]} """
            self.FilterByExpression(expression_text, layer_name)
        elif re.fullmatch(r">\s*-?\d+(\.\d+)?\s*", filter_value):
            res = GetNemIntFloatText()
            expression_text = f'"{filed}" > {res[0]}'
            self.FilterByExpression(expression_text, layer_name)
        elif re.fullmatch(r"-?\d+(\.\d+)?\s*::\s*-?\d+(\.\d+)?", filter_value):
            res = GetNemIntFloatText()
            expression_text = f""" "{filed}" BETWEEN {res[0]} AND {res[1]}"""
            self.FilterByExpression(expression_text, layer_name)
        elif re.fullmatch(r"^(0[1-9]|[12][0-9]|3[01])\-(0[1-9]|1[0-2])\-\d{4}$", filter_value) :
            res = GetNemIntFloatText()
            expression_text = f""" "{filed}" = '{filter_value}' """
            self.FilterByExpression(expression_text, layer_name)
            print(expression_text)
        else:
            expression_text = f""" "{filed}" ILIKE '%{filter_value}%'"""
            print(expression_text)
            self.FilterByExpression(expression_text, layer_name)

        

    def showAttributes(self):
        layer_name = self.combo_box1.currentText()
        layer = QgsProject.instance().mapLayersByName(layer_name)[0]
        self.table.clear()
        fields = layer.fields()
        feature_count = layer.featureCount()
        self.table.setRowCount(feature_count)
        self.table.setColumnCount(len(fields))
        for i, field in enumerate(fields):
            if field.alias() == "":
                self.table.setHorizontalHeaderItem(i, QTableWidgetItem(field.name()))
            else:
                self.table.setHorizontalHeaderItem(i, QTableWidgetItem(field.alias()))
        for i, feature in enumerate(layer.getFeatures()):
            for j, field in enumerate(fields):
                if isinstance(feature[field.name()], QtCore.QDate):
                    self.table.setItem(i, j, QTableWidgetItem(feature[field.name()].toString("yyyy-MM-dd")))
                else:
                    self.table.setItem(i, j, QTableWidgetItem(str(feature[field.name()])))



# -- Основний клас плагіна -- #
class OpsAutomatorPlugin:
    # Конструктор основного класу плагіна (НЕ ЧІПАТИ)
    def __init__(self, iface: QgisInterface):
        self.iface = iface
        self.log = PlgLogger().log
        self.panel = InfoTablePanel(self.iface)
        self.table_window = None
        
        # translation
        # initialize the locale
        self.locale: str = QgsSettings().value("locale/userLocale", QLocale().name())[
            0:2
        ]
        locale_path: Path = (
            DIR_PLUGIN_ROOT / "resources" / "i18n" / f"{__title__.lower()}_{self.locale}.qm"
        )
        self.log(message=f"Translation: {self.locale}, {locale_path}", log_level=4)
        if locale_path.exists():
            self.translator = QTranslator()
            self.translator.load(str(locale_path.resolve()))
            QCoreApplication.installTranslator(self.translator)
        print(DIR_PLUGIN_ROOT)

    # Метод що додає елементи інтерфейсу
    def initGui(self):
        self.panel = InfoTablePanel(self.iface)
        self.dock_widget = QDockWidget("Інформаційна панель", self.iface.mainWindow())
        self.dock_widget.setWidget(self.panel)
        self.iface.addDockWidget(Qt.RightDockWidgetArea, self.dock_widget)

        #== settings page within the QGIS preferences menu
        self.options_factory = PlgOptionsFactory()
        self.iface.registerOptionsWidgetFactory(self.options_factory)

        #== Actions
        self.action_help = QAction(
            QgsApplication.getThemeIcon("mActionHelpContents.svg"),
            self.tr("Help"),
            self.iface.mainWindow(),
        )
        self.action_help.triggered.connect(
            partial(QDesktopServices.openUrl, QUrl(__uri_homepage__))
        )

        self.action_settings = QAction(
            QgsApplication.getThemeIcon("console/iconSettingsConsole.svg"),
            self.tr("Settings"),
            self.iface.mainWindow(),
        )
        self.action_settings.triggered.connect(
            lambda: self.iface.showOptionsDialog(
                currentPage="mOptionsPage{}".format(__title__)
            )
        )

        #== Меню (кнопки допомоги та справки)
        self.iface.addPluginToMenu(__title__, self.action_settings)
        self.iface.addPluginToMenu(__title__, self.action_help)

        #== Документація
        self.iface.pluginHelpMenu().addSeparator()
        self.action_help_plugin_menu_documentation = QAction(
            QIcon(str(__icon_path__)),
            f"{__title__} - Documentation",
            self.iface.mainWindow(),
        )
        self.action_help_plugin_menu_documentation.triggered.connect(
            partial(QDesktopServices.openUrl, QUrl(__uri_homepage__))
        )

        self.iface.pluginHelpMenu().addAction(
            self.action_help_plugin_menu_documentation
        )

        #== Створення тубара
        self.toolbarAutomator = self.iface.addToolBar("Automator Toolbar")
        #== Кнопка пошуку по кадастровому або ПІБ
        self.action_search = QAction(QIcon(':/plugins/ops_automator/img/search.svg'), "Пошук", self.iface.mainWindow())
        self.action_search .triggered.connect(self.SearchParcel)
        self.toolbarAutomator.addAction(self.action_search)
        #== Кнопка вставки першої частина кадастрового номеру
        self.action_pastekoatuu= QAction(QIcon(':/plugins/ops_automator/img/koatuu.svg'), "Вставка КОАТУУ", self.iface.mainWindow())
        self.action_pastekoatuu.triggered.connect(self.SetCadastrFirsPart)
        self.toolbarAutomator.addAction(self.action_pastekoatuu)
        #== Текстове поле для вводу ПІБ або кадастрового
        self.input_panel = QLineEdit()
        self.input_panel.setPlaceholderText("Кадастровий номер або ПІБ")
        self.input_panel.setFixedWidth(200)
        self.toolbarAutomator.addWidget(self.input_panel)
        self.input_panel.textChanged.connect(self.onTextChanged)
        #== Кнопка очищення поля введення ПІБ та кадастру
        self.action_clearEditLine = QAction(QIcon(':/plugins/ops_automator/img/clear.svg'), "Очистити поле", self.iface.mainWindow())
        self.action_clearEditLine.triggered.connect(self.ClearInputPanel)
        self.toolbarAutomator.addAction(self.action_clearEditLine)
        #== Кнопка активації інструменту розрахунку нормативоно-грошової оцінки
        self.action_MonetaryValue = QAction(QIcon(':/plugins/ops_automator/img/norm.svg'), "Нормативно-грошова оцінка", self.iface.mainWindow())
        self.action_MonetaryValue.triggered.connect(self.MonetaryValue)
        self.toolbarAutomator.addAction( self.action_MonetaryValue)
        #== Кнопка активації генератора макетів
        self.action_MaketMenu = QAction(QIcon(':/plugins/ops_automator/img/doc.svg'), "Створити макет", self.iface.mainWindow())
        self.action_MaketMenu.triggered.connect(self.MaketMenu)
        self.toolbarAutomator.addAction( self.action_MaketMenu)
        #== Кнопка активації фільтра
        self.action_Filter = QAction(QIcon(':/plugins/ops_automator/img/filter.svg'), "Фільтрувати", self.iface.mainWindow())
        self.action_Filter.triggered.connect(self.ShowWindowFilter)
        self.toolbarAutomator.addAction( self.action_Filter)

    def SetCadastrFirsPart(self):
        self.input_panel.setText('5624689500')
    
    def ClearInputPanel(self):
        self.input_panel.clear()

    def ShowWindowFilter(self):
        self.window = FilterDataManage()
        self.window.show()

    # -- Маска для введення кадастрового номера
    def onTextChanged(self, text):
        pattern_one = r'^\d{10}$'
        pattern_two = r'^\d{10}:\d{2}$'
        pattern_three = r'^\d{10}:\d{2}:\d{3}$'

        def FormatText(text):
            text += ":"
            self.input_panel.setText(text)
            self.input_panel.setCursorPosition(len(text))

        if re.match(pattern_one, text):
            FormatText(text)
        if re.match(pattern_two, text):
            FormatText(text)
        if re.match(pattern_three, text):
            FormatText(text)
        else: 
            pass

    # -- Перекладач плагіна 
    def tr(self, message: str) -> str:
        """Get the translation for a string using Qt translation API.

        :param message: string to be translated.
        :type message: str

        :returns: Translated version of message.
        :rtype: str
        """
        return QCoreApplication.translate(self.__class__.__name__, message)
    
    # -- Вивантаження плагіна
    def unload(self):
        """Cleans up when plugin is disabled/uninstalled."""
        self.iface.removeDockWidget(self.dock_widget)
        self.iface.removeToolBarIcon(self.action_search)
        self.iface.removeToolBarIcon(self.action_MonetaryValue)
        self.iface.removeToolBarIcon(self.action_MaketMenu)
        del self.toolbarAutomator
        # -- Clean up menu
        self.iface.removePluginMenu(__title__, self.action_help)
        self.iface.removePluginMenu(__title__, self.action_settings)
        # -- Clean up preferences panel in QGIS settings
        self.iface.unregisterOptionsWidgetFactory(self.options_factory)
        # remove from QGIS help/extensions menu
        if self.action_help_plugin_menu_documentation:
            self.iface.pluginHelpMenu().removeAction(
                self.action_help_plugin_menu_documentation
            )
        # remove actions
        del self.action_settings
        del self.action_help

    # -- Функція пошуку ділянок за кадастровим номером або ПІБ
    def SearchParcel(self):
        canvas = self.iface.mapCanvas()
        text = self.input_panel.text()
        layer = QgsProject.instance().mapLayersByName('parcel')[0]
        if not layer:
            print("Шар не знайдено")
            return []
        pattern = r'^\d{10}:\d{2}:\d{3}:\d{4}$'
        pattern_pib = r'^[А-ЯЁа-яёІіЇїЄєҐґ’\-]+\s[А-ЯЁа-яёІіЇїЄєҐґ’\-]+\s[А-ЯЁа-яёІіЇїЄєҐґ’\-]+$'
        if re.match(pattern, text ):
            request = QgsFeatureRequest().setFilterExpression(f'"кадастровий номер" = \'{text}\'')
            for feature in layer.getFeatures(request):
                    layer.selectByIds([feature.id()])
                    canvas.zoomToSelected(layer)
        else:
            if re.match(pattern_pib, text):

                request = QgsFeatureRequest().setFilterExpression(f'"назва власника" = \'{text}\'')
                par_get = layer.getFeatures(request)
                features = list(par_get)
                count = len(features)
                if count > 1:
                    # Показуємо вікно з таблицею, якщо кількість об'єктів більше одного
                    self.ShowSearchTable(features)
                else:
                    for feature in layer.getFeatures(request):
                        print(feature)
                        layer.selectByIds([feature.id()])
                        canvas.zoomToSelected(layer)
            else:
                pass


    # -- Метод активації выкна з таблицею знайдених власників за ПІБ         
    def ShowSearchTable(self, features):
        if not self.table_window:
            self.table_window = InfoSerchTable()
        else:
            self.table_window.close()

        self.table_window.table_widget.setRowCount(len(features))

        for idx, feature in enumerate(features):
            self.table_window.table_widget.setItem(idx, 0, QTableWidgetItem(feature['назва власника']))
            self.table_window.table_widget.setItem(idx, 1, QTableWidgetItem(feature['кадастровий номер']))

        self.table_window.show()

    # -- Нормативоно грошова оцінка  
    def MonetaryValue(self):            
        def calculate_intersection_area(layer_1, layer_2):
            selected_features_layer_1 = layer_1.selectedFeatures()
            SelCount= layer_1.selectedFeatureCount()
            if not selected_features_layer_1:
                    print("У першому шарі не вибрано об'єктів.")                
            else: 
                if SelCount == 1:    
                    selected_feature_1 = selected_features_layer_1[0]
                    geom_1 = selected_feature_1.geometry()
                    features_layer_2 = layer_2.getFeatures()
                    intersection_areas = []
                    total_intersection_area = 0
                    for idx, feature_2 in enumerate(features_layer_2):
                        geom_2 = feature_2.geometry()
                        if geom_1.intersects(geom_2):
                            attributes = feature_2.attributes()
                            intersection_area = geom_1.intersection(geom_2).area()
                            intersection_areas.append([idx, attributes[2],intersection_area])
                            total_intersection_area += intersection_area
                    if abs(total_intersection_area - geom_1.area()) > 0.001:  # Перевірка на рівність з точністю 0.001
                        intersection_areas = None
                    return intersection_areas
                else:
                     pass
                
        # -- Клас результатів нормативно грошової оцінки (варто перенести позп основний клас) --#     
        class WindowResultNorm(QDialog):
            def __init__(self, iface, canvas):
                self.iface = iface
                self.canvas = canvas
                super().__init__()
                layout = QVBoxLayout()
                self.setWindowTitle("Розрахунок нормативно-грошової оцінки")
                self.label = QLabel("Вартість за одиницю площі")
                self.line_edit = QLineEdit()
                self.text_edit = QTextEdit()   
                self.button = QPushButton("Розрахунок")
                layout.addWidget(self.text_edit)
                layout.addWidget(self.label)
                layout.addWidget(self.line_edit)
                layout.addWidget(self.button)
                self.setLayout(layout)
                self.line_edit.setText("87")
                self.button.clicked.connect(self.RenderReport)
            
            # Рендер звіту розрахунків
            def RenderReport(self):
                your_layer_1 = QgsProject.instance().mapLayersByName('parcel')[0] 
                your_layer_2 = QgsProject.instance().mapLayersByName('Оціночні райони')[0] 
                
                rows = calculate_intersection_area(your_layer_1, your_layer_2)
                if rows == None:
                    self.text_edit.append("Сума площі перетину не дорівнює площі ділянкм")
                    pass
                else:
                    self.text_edit.append("ID, Комплексний коефіцієнт, Площа")
                    result_list = []
                    print(rows)
                    for line in rows:
                        self.text_edit.append(f"{line[0]},{line[1]},{line[2]}")
                        res = line[2] * float(self.line_edit.text()) * line[1]
                        result_list.append(res)
                    sum_res = sum(result_list)
                    self.text_edit.append("--------Результат--------")
                    self.text_edit.append(f"Вартість ділянки - {sum_res} грн.")
                    
        canvas = self.iface.mapCanvas()
        dialog = WindowResultNorm(self.iface, canvas)
        dialog.exec_()

    # Заповнення макетів документів
    def MaketMenu(self):
        # -- Клас діалогового вікна вибору макета--#
        class PluginDialog(QDialog):
            def __init__(self, canvas, iface):
                self.iface = iface
                self.canvas = canvas
                super().__init__()
                self.setWindowTitle("Пошук земельної ділянки")
                layout = QVBoxLayout()
                self.copy_parcel = QPushButton("Викопіювання з кадастру")
                self.zone_parcel = QPushButton("Витяг з плану зонування")
                self.copy_parcel.clicked.connect(self.GenerateParcelCopy)
                self.zone_parcel.clicked.connect(self.GenerateParcelInZone)
                layout.addWidget(self.copy_parcel)
                layout.addWidget(self.zone_parcel)
                self.setLayout(layout)

            # Очистка тимчасового шару від об'єктів
            def ClearTempLyr(self):
                layer = QgsProject.instance().mapLayersByName('temp_parcel')[0]
                if layer:
                    layer.startEditing() 
                    layer.selectAll()
                    layer.deleteSelectedFeatures() 
                    layer.commitChanges() 

            # Вставка ділянки на тимчасовий шар
            def PasteParcelTempLyr(self):
                    target_layer1 = QgsProject.instance().mapLayersByName('parcel')[0] 
                    selected_features = target_layer1.selectedFeatures()
                    target_layer = QgsProject.instance().mapLayersByName('temp_parcel')[0] 
                    if selected_features and selected_features:
                        target_layer.startEditing()
                        for feature in selected_features:
                            target_layer.addFeature(feature)
                        target_layer.commitChanges()
                    else:
                        print('Помилка не вибрано ділянку')

            # Отримання списку атрибутів шару
            def getAttributesByAttributeName(self):
                layer = QgsProject.instance().mapLayersByName('parcel')[0] 
                attributes = {}
                if layer is not None:
                    selected_features = layer.selectedFeatures()
                    if selected_features:
                        for feature in selected_features:
                            feature_attributes = {}
                            for field in layer.fields():
                                feature_attributes[field.name()] = feature[field.name()]
                            attributes = feature_attributes
                    else:
                        print("На даному шарі немає вибраних об'єктів")
                else:
                    print("Шар не обраний")
                return attributes
            
            # Переведення ПІБ у родовий відмінок
            def TransfomRodName(self, name):
                def inflect_name(name, inflection_case):
                    morph = pymorphy2.MorphAnalyzer(lang='uk')
                    name_parts = name.split()
                    inflected_parts = []
                    for part in name_parts:
                        parsed = morph.parse(part)[0]
                        inflected_part = parsed.inflect({inflection_case}).word
                        inflected_parts.append(inflected_part.capitalize())
                    return ' '.join(inflected_parts)
                inflected_name = inflect_name(name, 'gent')
                return inflected_name
            
            # Встановлення карти макету
            def SetCanvas(self, project, layout, id):
                map_item = layout.itemById('map')
                selected_layer = project.mapLayersByName('parcel')[0]
                selected_extent = selected_layer.boundingBoxOfSelected()
                selected_width = selected_extent.width()
                selected_height = selected_extent.height()
                new_extent = QgsRectangle(selected_extent.center().x() - selected_width / 2,
                                        selected_extent.center().y() - selected_height / 2,
                                        selected_extent.center().x() + selected_width / 2,
                                        selected_extent.center().y() + selected_height / 2)
                map_item.setExtent(new_extent)
                map_width = 15.213
                map_height = 9.874
                map_item.attemptResize(QgsLayoutSize(map_width, map_height, QgsUnitTypes.LayoutCentimeters))
                map_item.setScale(2000)
                layout.refresh()

            # Генератор викопіювання
            def GenerateParcelCopy(self):
                self.ClearTempLyr()
                self.PasteParcelTempLyr()
                project = QgsProject.instance()
                layout_manager = project.layoutManager()
                layout = layout_manager.layoutByName('Викопіювання')
                if layout is not None:
                    label_cadnum = layout.itemById('cadnum') 
                    label_parcelInfo = layout.itemById('info') 
                    atr = self.getAttributesByAttributeName()
                    InputTextInfoParsel = f"— земельна ділянка гр. {self.TransfomRodName(atr['owner'])} площею {atr['area_l']} га. {atr['use']}"
                    label_cadnum.setText(atr['cad_num'])
                    label_parcelInfo.setText(InputTextInfoParsel)
                    self.SetCanvas(project, layout, 'map')
                else:
                    print("Макет не знайдено")

            # Генератор витягу з плану зонування
            def GenerateParcelInZone(self):
                self.ClearTempLyr()
                self.PasteParcelTempLyr()
                project = QgsProject.instance()
                layout_manager = project.layoutManager()
                layout = layout_manager.layoutByName('Витяг')
                label_1 = layout.itemById('TEXT1') 

                if layout is not None:
                    atr = self.getAttributesByAttributeName()

                    text_1 = f"""
1. Територія, на яку розроблено містобудівну документацію: - земельна ділянка площею {atr['area_l']} га, яка пропонується до відведення у разі зміни її цільового призначення для будівництва та обслуговування житлового будинку, господарських будівель та споруд гр. {self.TransfomRodName(atr['owner'])} в с.Бармаки на території Шпанівської сільської ради (кадастровий номер {atr['cad_num']}). 
"""
                    label_1.setText(text_1)
                    self.SetCanvas(project, layout, 'map')
                else:
                    print("Макет не знайдено")

        canvas = self.iface.mapCanvas()
        dialog = PluginDialog(self.iface, canvas)
        dialog.exec_()

        
        

    