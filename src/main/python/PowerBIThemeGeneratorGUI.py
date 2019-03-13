import json
import sys
import logging
from functools import partial
from PyQt5 import QtWidgets
from PyQt5.QtGui import QColor, QFont
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QMainWindow, QAction, QFileDialog, QLabel, QPushButton, QColorDialog, \
    QDialogButtonBox
from PyQt5.QtWidgets import QStatusBar, QListWidget, QVBoxLayout, QHBoxLayout, QWidget, QTabWidget, QTabBar
from PyQt5.QtWidgets import QGroupBox, QCheckBox, QTreeWidget, QMessageBox, QLineEdit, QFormLayout, QDialog
from PowerBIThemeGenerator import PowerBIThemeGenerator
from CustomWidgets import CQListWidgetItemReportPages, CQListWidgetItemReportPageVisuals, \
    CQTreeWidgetItemVisualProperties
from Util import ColorUtil
from ErrorLoggingService import LogException, ShowErrorDialog
import AppInfo


class PowerBIThemeGeneratorWindow(QMainWindow):

    _powerBIThemeGenerator = None
    _pbiFilePath = None
    _horizontalLayoutTabVisualsTop = None
    _horizontalLayoutWelcomeScreen = None
    _groupBoxSelectedVisualPropertiesTree = None
    _listWidgetReportPages = None
    _listWidgetReportPageVisuals = None
    _verticalLayoutVisualsPropertiesTab = None
    _verticalLayoutMainWindow = None
    _treeWidgetSelectedVisualProperties = None
    _tabWidgetMainWindow = None
    _tabGeneralProperties = None
    _generalProperties = {}

    def __init__(self):

        super().__init__()

        # Windows Settings
        self.setWindowTitle(AppInfo.Name + ' - ' + AppInfo.Version)
        self.setGeometry(100, 100, 960, 720)

        # Creating central widget element
        self.centralWidget = QWidget(self)
        self.setCentralWidget(self.centralWidget)

        # Creating status bar
        self.statusBar = QStatusBar(self)
        self.setStatusBar(self.statusBar)

        # Creating menu bar
        self._createMenuBar()

        # Creating main window layout and adding widgets to them
        self._verticalLayoutMainWindow = QVBoxLayout(self.centralWidget)
        self._verticalLayoutMainWindow.addWidget(self._tabWidgetMainWindow)

        # for testing
        # self.__testOpenFileMethod()
        # end for testing

        self._horizontalLayoutWelcomeScreen = QHBoxLayout()
        self._createTabs()

        self.show()

    def _createMenuBar(self):

        # Create menu bar
        menuBar = self.menuBar()

        # Create root menu such as File, Edit, Help etc.
        menuBarFile = menuBar.addMenu('&File')
        menuBarHelp = menuBar.addMenu('&Help')

        # Create actions for Menu
        actionOpenPowerBIFile = QAction('Open &Power BI File', self)
        actionOpenPowerBIFile.setShortcut('Ctrl+O')
        actionOpenPowerBIFile.setStatusTip('Import Power BI file from which you want to export the visual settings to '
                                           'create theme')

        # actionReloadPowerBIFile = QAction('&Reload Power BI File', self)
        # actionReloadPowerBIFile.setShortcut('Ctrl+R')
        # actionReloadPowerBIFile.setStatusTip('Click to reload Power BI file if there are changes made in Power BI
        # file')

        actionGenerateTheme = QAction('&Generate Theme', self)
        actionGenerateTheme.setShortcut('Ctrl+G')
        actionGenerateTheme.setStatusTip('Click to generate theme and save theme file')

        actionQuit = QAction('&Quit', self)
        actionQuit.setShortcut('Ctrl+Q')
        actionQuit.setStatusTip('Click to quit application')

        actionAbout = QAction('&About', self)
        actionAbout.setStatusTip('Click to see detail application information and useful links')

        # Add actions to menu
        menuBarFile.addAction(actionOpenPowerBIFile)
        # menuBarFile.addAction(actionReloadPowerBIFile)
        menuBarFile.addAction(actionGenerateTheme)
        menuBarFile.addSeparator()
        menuBarFile.addAction(actionQuit)

        menuBarHelp.addAction(actionAbout)

        # Add global events i.e. what will happen when option is selected from menu bar
        actionOpenPowerBIFile.triggered.connect(self._openPowerBIFileDialog)
        actionQuit.triggered.connect(self.close)
        actionGenerateTheme.triggered.connect(self.generateTheme)
        actionAbout.triggered.connect(self._showAboutDialog)

    def _openPowerBIFileDialog(self):
        try:

            directory = '/'
            if self._pbiFilePath is not None:
                directory = self._pbiFilePath.split('/')[0:-1]
                directory = "/".join(directory)

            self._pbiFilePath = QFileDialog.getOpenFileName(
                self,
                caption='Select Power BI File',
                directory=directory,
                filter='Power BI Files(*.pbix)'
            )[0]

            if self._pbiFilePath is not '' and not None:
                if self._pbiFilePath.split('/')[-1].split('.')[-1] == 'pbix':
                    self._powerBIThemeGenerator = PowerBIThemeGenerator(self._pbiFilePath)
                else:
                    raise ValueError('Invalid Power BI File')
                self._reportVisualData = self._powerBIThemeGenerator.modifiedDataStructure()
                self._populateTabVisualProperties()
        except Exception as e:
            ShowErrorDialog(LogException(e))

    def __testOpenFileMethod(self):
        self._pbiFilePath = 'G:/Power BI Reports/Theme Template.pbix'
        self._powerBIThemeGenerator = PowerBIThemeGenerator(self._pbiFilePath)
        self._reportVisualData = self._powerBIThemeGenerator.modifiedDataStructure()

    def _createTabs(self):
        # Creating tabs
        self._tabWidgetMainWindow = QTabWidget(self.centralWidget)

        self._tabGeneralProperties = QWidget()
        self._tabVisualProperties = QWidget()

        self._tabWidgetMainWindow.addTab(self._tabGeneralProperties, 'General Properties')
        self._tabWidgetMainWindow.addTab(self._tabVisualProperties, 'Visuals Properties')

        tabBarMainWindow: QTabBar = self._tabWidgetMainWindow.tabBar()
        tabBarMainWindow.setTabToolTip(0, 'Add optional general properties of the theme')
        tabBarMainWindow.setTabToolTip(1, 'All the Power BI File visuals related properties will be visible here')

        self._verticalLayoutVisualsPropertiesTab = QVBoxLayout(self._tabVisualProperties)
        self._verticalLayoutMainWindow.addWidget(self._tabWidgetMainWindow)

        self._populateTabGeneralProperties()
        self._populateTabVisualProperties()

    def _populateTabVisualProperties(self):

        if self._pbiFilePath is None:
            self._showWelcomeScreenTabVisualProperties()
            self._verticalLayoutVisualsPropertiesTab.addLayout(self._horizontalLayoutWelcomeScreen)
        else:
            self._clearTabVisualProperties()
            self._horizontalLayoutTabVisualsTop = QHBoxLayout()
            self._verticalLayoutVisualsPropertiesTab.insertLayout(0, self._horizontalLayoutTabVisualsTop)

            self._createReportPageList()
            # self._create_extra_options()

    def _populateTabGeneralProperties(self):

        try:
            dataColors = []

            def __showMessageBoxInvalidHexColor(hexColor):
                messageBoxInvalidColor = QMessageBox()
                messageBoxInvalidColor.setIcon(QMessageBox.Critical)
                messageBoxInvalidColor.setWindowTitle('Invalid Hex Color Code')
                messageBoxInvalidColor.setText('"' + hexColor + '"' + ' is not valid hex color code')
                messageBoxInvalidColor.exec_()

            def __validateDataColors():
                try:
                    dataColors = [dataColor.strip() for dataColor in lineEditDataColors.text().split(',')]
                    for dataColor in dataColors:
                        if ColorUtil.IsValidHexColor(dataColor) is False and dataColor is not '':
                            dataColors.remove(dataColor)
                            __showMessageBoxInvalidHexColor(dataColor)
                        elif dataColor is '':
                            dataColors.remove(dataColor)
                    lineEditDataColors.setText(','.join(dataColors))
                except Exception as e:
                    ShowErrorDialog(LogException(e))

            def __showColorPickerDialog():
                try:
                    sender = self.sender().objectName()
                    color: QColor = QColorDialog().getColor()
                    if color.isValid():
                        hexColor = color.name()
                        buttonBackgroundColor = 'background-color: ' + hexColor
                        if sender == 'push_button_data_color':
                            dataColors.append(hexColor)
                            self._generalProperties['dataColors'] = dataColors
                            lineEditDataColors.setText(','.join(dataColors))
                        elif sender == 'push_button_background_color':
                            pushButtonBackgroundColorPicker.setStyleSheet(buttonBackgroundColor)
                            lineEditBackgroundColor.setText(hexColor)
                            self._generalProperties['background'] = hexColor
                        elif sender == 'push_button_foreground_color':
                            pushButtonForegroundColorPicker.setStyleSheet(buttonBackgroundColor)
                            lineEditForegroundColor.setText(hexColor)
                            self._generalProperties['foreground'] = hexColor
                        elif sender == 'push_button_table_accent_color':
                            pushButtonTableAccentColorPicker.setStyleSheet(buttonBackgroundColor)
                            lineEditTableAccent.setText(hexColor)
                            self._generalProperties['tableAccent'] = hexColor
                except Exception as e:
                    ShowErrorDialog(LogException(e))

            def __lineEditEditingFinished():
                try:
                    sender = self.sender().objectName()
                    if sender == 'line_edit_data_color':
                        __validateDataColors()
                        nonlocal dataColors
                        dataColors = lineEditDataColors.text().split(",")
                        self._generalProperties['dataColors'] = dataColors
                        return
                    elif sender == 'line_edit_background_color':
                        textColor = lineEditBackgroundColor.text()
                    elif sender == 'line_edit_foreground_color':
                        textColor = lineEditForegroundColor.text()
                    elif sender == 'line_edit_table_accent_color':
                        textColor = lineEditTableAccent.text()
                    elif sender == 'line_edit_theme_name':
                        self._generalProperties['name'] = lineEditThemeName.text()
                        return
                    validColor = ColorUtil.IsValidHexColor(textColor)
                    if validColor is False and textColor is not '':
                        __showMessageBoxInvalidHexColor(textColor)
                        if sender == 'line_edit_background_color':
                            lineEditBackgroundColor.setText('')
                        elif sender == 'line_edit_foreground_color':
                            lineEditForegroundColor.setText('')
                        elif sender == 'line_edit_table_accent_color':
                            lineEditTableAccent.setText('')
                    else:
                        buttonBackgroundColor = 'background-color: ' + textColor
                        if sender == 'line_edit_background_color':
                            self._generalProperties['background'] = textColor
                            pushButtonBackgroundColorPicker.setStyleSheet(buttonBackgroundColor)
                        elif sender == 'line_edit_foreground_color':
                            pushButtonForegroundColorPicker.setStyleSheet(buttonBackgroundColor)
                            self._generalProperties['foreground'] = textColor
                        elif sender == 'line_edit_table_accent_color':
                            pushButtonTableAccentColorPicker.setStyleSheet(buttonBackgroundColor)
                            self._generalProperties['tableAccent'] = textColor

                except Exception as e:
                    ShowErrorDialog(LogException(e))

            labelThemeName = QLabel('Theme Name: ')
            lineEditThemeName = QLineEdit()
            lineEditThemeName.setObjectName('line_edit_theme_name')
            lineEditThemeName.editingFinished.connect(__lineEditEditingFinished)

            horizontalLayoutDataColor = QHBoxLayout()
            labelDataColors = QLabel('Data Colors: ')
            labelDataColors.setToolTip('These colors will be visible in color picker and will apply as data colors in '
                                       'charts')
            lineEditDataColors = QLineEdit()
            lineEditDataColors.editingFinished.connect(__lineEditEditingFinished)
            lineEditDataColors.setObjectName('line_edit_data_color')
            pushButtonDataColorPicker = QPushButton()
            pushButtonDataColorPicker.setToolTip('Click to pick colors for Data Colors')
            pushButtonDataColorPicker.setObjectName('push_button_data_color')
            pushButtonDataColorPicker.clicked.connect(__showColorPickerDialog)
            horizontalLayoutDataColor.addWidget(lineEditDataColors)
            horizontalLayoutDataColor.addWidget(pushButtonDataColorPicker)
            # TO DO: Add generate colors button

            horizontalLayoutBackgroundColor = QHBoxLayout()
            labelBackgroundColor = QLabel('Background Color: ')
            labelBackgroundColor.setToolTip('Applies to button fill and combo chart label background. How these colors '
                                            'are used depends on the specific visual style that\'s applied.')
            lineEditBackgroundColor = QLineEdit()
            lineEditBackgroundColor.setObjectName('line_edit_background_color')
            lineEditBackgroundColor.editingFinished.connect(__lineEditEditingFinished)
            pushButtonBackgroundColorPicker = QPushButton()
            pushButtonBackgroundColorPicker.setToolTip('Click to pick color for Background Color')
            pushButtonBackgroundColorPicker.setObjectName('push_button_background_color')
            pushButtonBackgroundColorPicker.clicked.connect(__showColorPickerDialog)
            horizontalLayoutBackgroundColor.addWidget(lineEditBackgroundColor)
            horizontalLayoutBackgroundColor.addWidget(pushButtonBackgroundColorPicker)

            horizontalLayoutForegroundColor = QHBoxLayout()
            labelForegroundColor = QLabel('Foreground Color: ')
            labelForegroundColor.setToolTip('Applies to textbox text, KPI goal text, multi-row card text, card value '
                                            'text, gauge callout text, vertical slicer element text, and table and '
                                            'matrix total and values text.')
            lineEditForegroundColor = QLineEdit()
            lineEditForegroundColor.setObjectName('line_edit_foreground_color')
            lineEditForegroundColor.editingFinished.connect(__lineEditEditingFinished)
            pushButtonForegroundColorPicker = QPushButton()
            pushButtonForegroundColorPicker.setToolTip('Click to pick color for Foreground Color')
            pushButtonForegroundColorPicker.setObjectName('push_button_foreground_color')
            pushButtonForegroundColorPicker.clicked.connect(__showColorPickerDialog)
            horizontalLayoutForegroundColor.addWidget(lineEditForegroundColor)
            horizontalLayoutForegroundColor.addWidget(pushButtonForegroundColorPicker)

            horizontalLayoutTableAccent = QHBoxLayout()
            labelTableAccent = QLabel('Table Accent: ')
            labelTableAccent.setToolTip('Table and matrix visuals apply these styles by default.')
            lineEditTableAccent = QLineEdit()
            lineEditTableAccent.setObjectName('line_edit_table_accent_color')
            lineEditTableAccent.editingFinished.connect(__lineEditEditingFinished)
            pushButtonTableAccentColorPicker = QPushButton()
            pushButtonTableAccentColorPicker.setToolTip('Click to pick color for Table Accent Color')
            pushButtonTableAccentColorPicker.setObjectName('push_button_table_accent_color')
            pushButtonTableAccentColorPicker.clicked.connect(__showColorPickerDialog)
            horizontalLayoutTableAccent.addWidget(lineEditTableAccent)
            horizontalLayoutTableAccent.addWidget(pushButtonTableAccentColorPicker)

            formLayoutGeneralProperties = QFormLayout(self._tabGeneralProperties)
            formLayoutGeneralProperties.addRow(labelThemeName, lineEditThemeName)
            formLayoutGeneralProperties.addRow(labelDataColors, horizontalLayoutDataColor)
            formLayoutGeneralProperties.addRow(labelBackgroundColor, horizontalLayoutBackgroundColor)
            formLayoutGeneralProperties.addRow(labelForegroundColor, horizontalLayoutForegroundColor)
            formLayoutGeneralProperties.addRow(labelTableAccent, horizontalLayoutTableAccent)
        except Exception as e:
            ShowErrorDialog(LogException(e))

    def _showWelcomeScreenTabVisualProperties(self):
        welcomeLabel = QLabel('Go to File menu or press Ctrl+O to select Power BI File')
        self._horizontalLayoutWelcomeScreen.addStretch()
        self._horizontalLayoutWelcomeScreen.addWidget(welcomeLabel)
        self._horizontalLayoutWelcomeScreen.addStretch()

    def _removeWelcomeScreenFromTabVisualProperties(self):
        self._deleteLayout(self._horizontalLayoutWelcomeScreen)

    def _clearTabVisualProperties(self):
        self._removeWelcomeScreenFromTabVisualProperties()
        self._deleteLayout(self._verticalLayoutVisualsPropertiesTab)
        self._listWidgetReportPages = None
        self._listWidgetReportPageVisuals = None
        self._treeWidgetSelectedVisualProperties = None

    def _deleteLayout(self, layout):
        if layout is not None:
            while layout.count():
                item = layout.takeAt(0)
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()
                else:
                    self._deleteLayout(item.layout())
            # sip.delete(layout)

    def _createReportPageList(self):
        try:
            def __reportPageListValueSelected(item):
                try:
                    self._createReportPageVisualsList(item.GetMetaData()['reportPageSection'])
                except Exception as e:
                    ShowErrorDialog(LogException(e))

            if self._listWidgetReportPages is not None:
                self._listWidgetReportPages.clear()

            groupBoxReportPageList = QGroupBox(self.centralWidget)
            groupBoxReportPageList.setTitle("Report Pages")
            verticalLayoutReportPageList = QVBoxLayout(groupBoxReportPageList)

            self._listWidgetReportPages = QListWidget()

            i = 0
            for reportSection in self._powerBIThemeGenerator.get_report_sections():
                reportListWidgetItem = CQListWidgetItemReportPages(str(reportSection['displayName']))
                reportListWidgetItem.SetMetaData(reportSection)
                reportListWidgetItem.setToolTip(reportSection['name'])
                self._listWidgetReportPages.insertItem(i, reportListWidgetItem)
                i += 1

            verticalLayoutReportPageList.addWidget(self._listWidgetReportPages)
            self._horizontalLayoutTabVisualsTop.addWidget(groupBoxReportPageList)
            # __reportPageListValueSelected(self._listWidgetReportPages.item(0))

            self._listWidgetReportPages.itemClicked.connect(__reportPageListValueSelected)

        except Exception as e:
            ShowErrorDialog(LogException(e))

    def _createReportPageVisualsList(self, reportPageSection=None):

        try:
            # print(self._reportVisualData[reportPageSection])
            # print('inside report page visual list', reportPageSection)

            def __updateStateInDataStructure(item):
                vizProperties = item.GetVisualProperties()
                if item.checkState() == Qt.Checked:
                    (self._reportVisualData[vizProperties['report_section']]
                        .get('visuals')[vizProperties['visual_index']]['__selected']) = True
                else:
                    (self._reportVisualData[vizProperties['report_section']]
                        .get('visuals')[vizProperties['visual_index']]['__selected']) = False

            def __getVisualProperties(item: CQListWidgetItemReportPageVisuals):
                try:
                    __updateStateInDataStructure(item)
                    self._createSelectedVisualPropertiesTree(item.GetVisualProperties())
                except Exception as e:
                    ShowErrorDialog(LogException(e))

            def __selectDeselectAll():
                try:
                    if self._listWidgetReportPageVisuals is not None:
                        for i in range(self._listWidgetReportPageVisuals.count()):
                            item = self._listWidgetReportPageVisuals.item(i)
                            if self.sender().objectName() == 'button_select_all':
                                item.setCheckState(Qt.Checked)
                            else:
                                item.setCheckState(Qt.Unchecked)
                            __updateStateInDataStructure(item)
                except Exception as e:
                    ShowErrorDialog(LogException(e))

            groupBoxReportPageVisualsList = QGroupBox(self.centralWidget)
            groupBoxReportPageVisualsList.setTitle("Visuals in Selected Report Page")
            # groupBoxReportPageVisualsList.setToolTip("All visuals present in the selected page will be visible below")
            verticalLayoutReportPageVisualsList = QVBoxLayout(groupBoxReportPageVisualsList)

            horizontalLayoutSelectDeselectButton = QHBoxLayout()
            pushButtonSelectAll = QPushButton('Select All')
            pushButtonSelectAll.setObjectName('button_select_all')
            pushButtonDeselectAll = QPushButton('Deselect All')
            pushButtonDeselectAll.setObjectName('button_deselect_all')
            horizontalLayoutSelectDeselectButton.addWidget(pushButtonSelectAll)
            horizontalLayoutSelectDeselectButton.addWidget(pushButtonDeselectAll)
            verticalLayoutReportPageVisualsList.addLayout(horizontalLayoutSelectDeselectButton)

            if self._listWidgetReportPageVisuals is None:
                self._listWidgetReportPageVisuals = QListWidget()
                verticalLayoutReportPageVisualsList.addWidget(self._listWidgetReportPageVisuals)
                self._horizontalLayoutTabVisualsTop.addWidget(groupBoxReportPageVisualsList)
                self._listWidgetReportPageVisuals.itemClicked.connect(__getVisualProperties, Qt.UniqueConnection)
            else:
                self._listWidgetReportPageVisuals.clear()

            i = 0
            for visualIndex, visual in enumerate(self._reportVisualData[reportPageSection].get('visuals')):
                visualType = visual['visual_type']
                visualTitle = visual.get('objects').get('title')
                # print(visualType, visualTitle)
                if visualTitle is None or visualTitle.get('text') is None:
                    listDisplayText = str(visualType)
                else:
                    listDisplayText = str(visualType) + " - " + visualTitle.get('text')
                pageVisualsList = CQListWidgetItemReportPageVisuals(listDisplayText)
                pageVisualsList.setFlags(pageVisualsList.flags() | Qt.ItemIsUserCheckable)
                if visual['__selected'] is False:
                    pageVisualsList.setCheckState(Qt.Unchecked)
                else:
                    pageVisualsList.setCheckState(Qt.Checked)
                pageVisualsList.SetVisualProperties(visual, reportPageSection, visualIndex)
                # print(pageVisualsList.GetVisualProperties())
                self._listWidgetReportPageVisuals.insertItem(i, pageVisualsList)
                i += 1

            pushButtonSelectAll.clicked.connect(__selectDeselectAll)
            pushButtonDeselectAll.clicked.connect(__selectDeselectAll)

            # self._listWidgetReportPageVisuals.setCurrentRow(0)
            # __getVisualProperties(self._listWidgetReportPageVisuals.item(0))

        except Exception as e:
            ShowErrorDialog(LogException(e))

    def _createSelectedVisualPropertiesTree(self, visualProperties=None):
        try:

            def __getVisualProperty(item: CQTreeWidgetItemVisualProperties):
                try:
                    visualProperties = item.GetVisualProperties()
                    if visualProperties is not None:
                        visualProperty = (self._reportVisualData[visualProperties['report_section']]
                                          .get('visuals')[visualProperties['visual_index']].get('objects')
                                          .get(str(item.data(0, 0))))
                        return visualProperty
                except Exception as e:
                    ShowErrorDialog(LogException(e))

            def __updateStateInDataStructure(item: CQTreeWidgetItemVisualProperties):
                try:
                    visualProperty = __getVisualProperty(item)
                    if visualProperty is not None:
                        if item.checkState(0) == Qt.Unchecked:
                            visualProperty['__selected'] = False
                        else:
                            visualProperty['__selected'] = True
                except Exception as e:
                    ShowErrorDialog(LogException(e))

            def __updateWildCardState(state, item=None):
                try:
                    visualProperty = __getVisualProperty(item)
                    if visualProperty is not None:
                        if state == Qt.Checked:
                            visualProperty['__wildcard'] = True
                            # print(visualProperty)
                        elif visualProperty.get('__wildcard') is not None and state == Qt.Unchecked:
                            visualProperty.pop('__wildcard')
                except Exception as e:
                    ShowErrorDialog(LogException(e))

            def __treeWidgetVisualsPropertiesClicked(item: CQTreeWidgetItemVisualProperties):
                try:
                    __updateStateInDataStructure(item)
                except Exception as e:
                    ShowErrorDialog(LogException(e))

            def __selectDeselectAll():
                try:
                    if self._treeWidgetSelectedVisualProperties is not None:
                        for i in range(self._treeWidgetSelectedVisualProperties.topLevelItemCount()):
                            item = self._treeWidgetSelectedVisualProperties.topLevelItem(i)
                            if self.sender().objectName() == 'button_select_all':
                                item.setCheckState(0, Qt.Checked)
                            else:
                                item.setCheckState(0, Qt.Unchecked)
                            __updateStateInDataStructure(item)
                except Exception as e:
                    ShowErrorDialog(LogException(e))

            self._groupBoxSelectedVisualPropertiesTree = QGroupBox(self.centralWidget)
            self._groupBoxSelectedVisualPropertiesTree.setTitle("Selected Visual Properties")
            verticalLayoutSelectedVisualPropertiesTree = QVBoxLayout(self._groupBoxSelectedVisualPropertiesTree)

            horizontalLayoutSelectDeselectButton = QHBoxLayout()
            pushButtonSelectAll = QPushButton('Select All')
            pushButtonSelectAll.setObjectName('button_select_all')
            pushButtonDeselectAll = QPushButton('Deselect All')
            pushButtonDeselectAll.setObjectName('button_deselect_all')
            horizontalLayoutSelectDeselectButton.addWidget(pushButtonSelectAll)
            horizontalLayoutSelectDeselectButton.addWidget(pushButtonDeselectAll)
            verticalLayoutSelectedVisualPropertiesTree.addLayout(horizontalLayoutSelectDeselectButton)

            if self._treeWidgetSelectedVisualProperties is None:
                self._treeWidgetSelectedVisualProperties = QTreeWidget()
                verticalLayoutSelectedVisualPropertiesTree.addWidget(self._treeWidgetSelectedVisualProperties)
                self._verticalLayoutVisualsPropertiesTab.insertWidget(1, self._groupBoxSelectedVisualPropertiesTree)
                self._treeWidgetSelectedVisualProperties.itemClicked.connect(__treeWidgetVisualsPropertiesClicked)
            else:
                self._treeWidgetSelectedVisualProperties.clear()

            self._treeWidgetSelectedVisualProperties.setHeaderLabels(['Property', 'Value', 'Wild Card'])
            treeHeader = self._treeWidgetSelectedVisualProperties.header()
            treeHeader.setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)
            # treeHeader.setStretchLastSection(False)

            visualObjects = visualProperties['visual_properties'].get('objects')
            for object, objectValues in visualObjects.items():
                if len(objectValues.keys()) > 1:  # greater than 1 because one default key is __selected
                    parent = CQTreeWidgetItemVisualProperties(self._treeWidgetSelectedVisualProperties, [str(object)])
                    parent.SetVisualProperties(visualProperties)
                    wildCardCheckBox = QCheckBox()
                    wildCardCheckBox.stateChanged.connect(partial(__updateWildCardState, item=parent))
                    self._treeWidgetSelectedVisualProperties.setItemWidget(parent, 2, wildCardCheckBox)
                    if objectValues['__selected'] is False:
                        parent.setCheckState(0, Qt.Unchecked)
                    else:
                        parent.setCheckState(0, Qt.Checked)
                    for objectKey, objectValue in objectValues.items():
                        if objectKey is not '__selected':
                            mObjectV = None
                            if type(objectValue) is dict and objectValue.get('solid') is not None:
                                mObjectV = objectValue.get('solid').get('color')
                            else:
                                mObjectV = objectValue
                            CQTreeWidgetItemVisualProperties(parent, [str(objectKey), str(mObjectV)])
                else:
                    objectValues['__selected'] = False

            pushButtonSelectAll.clicked.connect(__selectDeselectAll)
            pushButtonDeselectAll.clicked.connect(__selectDeselectAll)

        except Exception as e:
            ShowErrorDialog(LogException(e))

    def generateTheme(self):
        try:
            correctVisualData = True
            correctWildcardData = True
            themeName = self._generalProperties.get('name')
            themeData = {
                'name': themeName
                if themeName is not None and len(themeName.strip()) > 0 else 'My Theme',
            }
            if self._generalProperties.get('dataColors') is not None:
                dataColors = list(filter(None, self._generalProperties.get('dataColors')))

            else:
                dataColors = []
            background = self._generalProperties.get('background')
            foreground = self._generalProperties.get('foreground')
            tableAccent = self._generalProperties.get('tableAccent')

            if dataColors is not None and dataColors is not "":
                themeData['dataColors'] = dataColors
            if background is not None and background is not "":
                themeData['background'] = background
            if foreground is not None and foreground is not "":
                themeData['foreground'] = foreground
            if tableAccent is not None and tableAccent is not "":
                themeData['tableAccent'] = tableAccent

            selectedVisualsList = [[], [], []]  # page name, visual type and visual data or visual objects
            wildCardPropertiesList = [[], [], []]  # page name, property type and property data
            visualObjectIndex = 0
            themeData['visualStyles'] = {}
            if self._pbiFilePath is not None:
                for reportPage in self._reportVisualData:
                    for visual in self._reportVisualData[reportPage].get('visuals'):
                        pageName = self._reportVisualData[reportPage]['reportPageDisplayName']
                        if visual['__selected'] is True:
                            selectedVisualsList[0].append(pageName)
                            selectedVisualsList[1].append(visual['visual_type'])
                            selectedVisualsList[2].append({})
                            visualObjects = visual.get('objects')
                            for object in visualObjects:
                                if visualObjects[object]['__selected'] is True:
                                    selectedVisualsList[2][visualObjectIndex][object] = [visualObjects[object]]
                                    if visualObjects[object].get('__wildcard') is not None:
                                        # wildCardObjects[object] = [visualObjects[object]]
                                        wildCardPropertiesList[0].append(pageName)
                                        wildCardPropertiesList[1].append(object)
                                        wildCardPropertiesList[2].append(visualObjects[object])

                            visualObjectIndex += 1
                if len(selectedVisualsList[1]) > 0:
                    if len(selectedVisualsList[1]) != len(set(selectedVisualsList[1])):
                        message = 'Multiple visual of same type selected from different report pages. See details below.' \
                                  ' Please select one visual type for only once.'
                        detailMessage = ''
                        seen = set()
                        repeatedVisuals = []
                        for selectedVisual in selectedVisualsList[1]:
                            if selectedVisual in seen:
                                repeatedVisuals.append(selectedVisual)
                            else:
                                seen.add(selectedVisual)

                        for i, selectedVisual in enumerate(selectedVisualsList[1]):
                            if selectedVisual in repeatedVisuals:
                                detailMessage += selectedVisual + ' is selected in report page'\
                                                 + selectedVisualsList[0][i] + '\n'
                        correctVisualData = False
                        self._showDialogAfterThemeGeneration(message, detailMessage=detailMessage)

                    else:
                        # adding wild card data
                        if len(wildCardPropertiesList[1]) > 0:
                            if len(wildCardPropertiesList[1]) != len(set(wildCardPropertiesList[1])):
                                message = 'Multiple wild card properties of same types are selected. Please only select '\
                                          'unique properties for wild card from all the visuals. As a result this' \
                                          'properties will not be added as wild card property in them file. ' \
                                          'See details below'
                                detailMessage = ''
                                seen = set()
                                repeatedWildCardProperties = []
                                for wildCardProperty in wildCardPropertiesList[1]:
                                    if wildCardProperty in seen:
                                        repeatedWildCardProperties.append(wildCardProperty)
                                    else:
                                        seen.add(wildCardProperty)

                                for i, wildCardProperty in enumerate(wildCardPropertiesList[1]):
                                    if wildCardProperty in repeatedWildCardProperties:
                                        detailMessage += wildCardProperty + 'is selected in report page' +\
                                                         wildCardPropertiesList[0][i]
                                correctWildcardData = False
                                self._showDialogAfterThemeGeneration(message, detailMessage=detailMessage)
                            else:
                                themeData['visualStyles']['*'] = {
                                    '*': {}
                                }
                                for i in range(len(wildCardPropertiesList[1])):
                                    themeData['visualStyles']['*']['*'][wildCardPropertiesList[1][i]] = [wildCardPropertiesList[2][i]]

                        # adding visual data
                        for i in range(len(selectedVisualsList[1])):
                            themeData['visualStyles'][selectedVisualsList[1][i]] = {
                                "*": selectedVisualsList[2][i]
                            }
            if correctVisualData is True and correctWildcardData is True:
                self._saveThemeFile(themeData, 'C:/Users/bjadav/Desktop/', themeData['name'])

        except Exception as e:
            ShowErrorDialog(LogException(e))

    def _showDialogAfterThemeGeneration(self, message, messageIcon=QMessageBox.Information, detailMessage=None):
        try:
            messageBoxThemeGeneration = QMessageBox()
            messageBoxThemeGeneration.setIcon(messageIcon)
            messageBoxThemeGeneration.setText(message)
            if detailMessage is not None:
                messageBoxThemeGeneration.setInformativeText(detailMessage)
            messageBoxThemeGeneration.setWindowTitle("Theme Generation Information")
            messageBoxThemeGeneration.setStandardButtons(QMessageBox.Ok)

            messageBoxThemeGeneration.exec_()
        except Exception as e:
            ShowErrorDialog(LogException(e))

    def _saveThemeFile(self, themeData, saveFileDirectory, fileName):
        try:
            _themeData = json.loads(json.dumps(themeData))
            # Removing internal keys such as __selected from final output
            for visualType, visualProperties in _themeData['visualStyles'].items():
                for propertyType, propertyValue in visualProperties['*'].items():
                    if propertyValue[0].get('__selected') is not None:
                        propertyValue[0].pop('__selected')
                    if propertyValue[0].get('__wildcard') is not None:
                        propertyValue[0].pop('__wildcard')

            initialSaveFilePath = saveFileDirectory + fileName + '.json'
            saveFile = QFileDialog.getSaveFileName(self, 'Save Theme File', initialSaveFilePath,
                                                   filter='JSON file(*.json)')[0]
            if saveFile is not '':
                with open(saveFile, 'w') as themeFile:
                    json.dump(_themeData, themeFile, indent=4)
                message = 'Successfully generated theme'
                self._showDialogAfterThemeGeneration(message)
        except Exception as e:
            ShowErrorDialog(LogException(e))

    def _showAboutDialog(self):
        try:
            dialogAbout = QDialog()
            dialogAbout.resize(320, 180)
            dialogButtonBoxAbout = QDialogButtonBox(dialogAbout)
            dialogButtonBoxAbout.setLayoutDirection(Qt.LeftToRight)
            dialogButtonBoxAbout.setAutoFillBackground(False)
            dialogButtonBoxAbout.setOrientation(Qt.Horizontal)
            dialogButtonBoxAbout.setStandardButtons(QtWidgets.QDialogButtonBox.Ok)
            dialogButtonBoxAbout.setCenterButtons(True)
            dialogButtonBoxAbout.accepted.connect(dialogAbout.accept)

            verticalLayoutAboutDialog = QVBoxLayout(dialogAbout)

            dialogAbout.setWindowTitle('About ' + AppInfo.Name)

            labelPowerBIThemeGeneratorFont = QFont()
            labelPowerBIThemeGeneratorFont.setPointSize(16)
            labelPowerBIThemeGeneratorFont.setFamily('Calibri')
            labelPowerBIThemeGenerator = QLabel(AppInfo.Name)
            labelPowerBIThemeGenerator.setFont(labelPowerBIThemeGeneratorFont)

            labelVersionFont = QFont()
            labelVersionFont.setPointSize(8)
            labelVersion = QLabel(AppInfo.Version)
            labelVersion.setContentsMargins(0, 0, 0, 0)

            labelCreatedBy = QLabel('Created By ' + AppInfo.Author)
            link = '<a href="' + AppInfo.GitHubRepoIssuesURL + '">Click here to report bugs</a>'
            labelBugs = QLabel(link)
            labelBugs.setOpenExternalLinks(True)

            verticalLayoutAboutDialog.addWidget(labelPowerBIThemeGenerator)
            verticalLayoutAboutDialog.addWidget(labelVersion)
            verticalLayoutAboutDialog.addWidget(labelCreatedBy)
            verticalLayoutAboutDialog.addWidget(labelBugs)
            verticalLayoutAboutDialog.addStretch()
            verticalLayoutAboutDialog.addWidget(dialogButtonBoxAbout)
            dialogAbout.exec_()
        except Exception as e:
            ShowErrorDialog(LogException(e))

    # def _create_extra_options(self):
    #     groupBoxExtraOptions = QGroupBox(self.centralWidget)
    #     groupBoxExtraOptions.setTitle('Extra Options')
    #     horizontalLayoutExtraOptions = QHBoxLayout(groupBoxExtraOptions)
    #     checkBoxCombineColors = QCheckBox('Combine Colors')
    #     horizontalLayoutExtraOptions.addWidget(checkBoxCombineColors)
    #     self._verticalLayoutVisualsPropertiesTab.addWidget(groupBoxExtraOptions)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    try:
        window = PowerBIThemeGeneratorWindow()
    except Exception as e:
        logging.exception("Error")
    sys.exit(app.exec_())
