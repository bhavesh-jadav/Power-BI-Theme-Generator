from PyQt5.QtWidgets import QListWidgetItem, QTreeWidgetItem, QCheckBox

class CQListWidgetItemReportPages(QListWidgetItem):

    _metaData = None

    def __init__(self, data=None):
        super().__init__(data)

    def SetMetaData(self, reportSection):
        self._metaData = {
            'reportPageSection': reportSection['name']
        }

    def GetMetaData(self):
        return self._metaData


class CQListWidgetItemReportPageVisuals(QListWidgetItem):

    _visualProperties = None

    def __init__(self, data=None):
        super().__init__(data)

    def SetVisualProperties(self, visual_properties, reportSection, visualIndex):
        self._visualProperties = {
            'visual_properties': visual_properties,
            'report_section': reportSection,
            'visual_index': visualIndex
        }

    def GetVisualProperties(self):
        return self._visualProperties


class CQTreeWidgetItemVisualProperties(QTreeWidgetItem):

    _visualProperties = None

    def __init__(self, *args):
        super().__init__(*args)

    def SetVisualProperties(self, visual_properties):
        self._visualProperties = visual_properties

    def GetVisualProperties(self):
        return self._visualProperties
