from zipfile import ZipFile, BadZipFile, LargeZipFile
from Util import ColorUtil
import json
import io
from ErrorLoggingService import LogException, ShowErrorDialog


class PowerBIThemeGenerator:
    _default_theme_accent_colors = ["#FFFFFF", "#000000", "#01B8AA", "#374649", "#FD625E", "#F2C80F", "#5F6B6D",
                                    "#8AD4EB", "#FE9666", "#A66999", "#3599B8", "#DFBFBF", "#4AC5BB", "#5F6B6D",
                                    "#FB8281", "#F4D25A", "#7F898A", "#A4DDEE", "#FDAB89", "#B687AC", "#28738A",
                                    "#A78F8F", "#168980", "#293537", "#BB4A4A", "#B59525", "#475052", "#6A9FB0",
                                    "#BD7150", "#7B4F71", "#1B4D5C", "#706060", "#0F5C55", "#1C2325", "#7D3231",
                                    "#796419", "#303637", "#476A75", "#7E4B36", "#52354C", "#0D262E", "#544848",
                                    ]

    def __init__(self, pbi_file_path=None):
        if pbi_file_path is None:
            raise ValueError('No power bi file provided')
        else:
            if pbi_file_path.endswith('.pbix'):
                self._pbi_file_path = pbi_file_path
            else:
                raise ValueError(".pbix file not provided")
        self._get_report_layout_data()

    def _get_report_layout_data(self):
        self._layout_data = None
        layout_path_in_zip = 'Report/Layout'
        try:
            with ZipFile(self._pbi_file_path, 'r') as power_bi_file:
                self._layout_data = json.load(io.BytesIO(power_bi_file.read(layout_path_in_zip)))
        except (BadZipFile, LargeZipFile) as e:
            ShowErrorDialog(LogException(e))

    def get_report_sections(self):
        return self._layout_data['sections']

    def _get_all_visuals_type_in_report(self):
        try:
            all_visual_types = []
            for section in self._layout_data['sections']:
                for visualContainer in section['visualContainers']:
                    visual_type = json.loads(visualContainer['config'])['singleVisual']['visualType']
                    if visual_type not in all_visual_types:
                        all_visual_types.append(visual_type)
            return all_visual_types
        except Exception as e:
            ShowErrorDialog(LogException(e))

    def _get_page_wise_visual_properties(self, report_page_section):

        try:
            page_visuals_properties = {
                'reportPageDisplayName': report_page_section['displayName'],
                'visuals': [],
            }

            for visualContainer in report_page_section['visualContainers']:
                config = json.loads(visualContainer['config'])
                vcObjects = config['singleVisual'].get('vcObjects')
                objects = config['singleVisual'].get('objects')

                # print(objects)

                if objects is None and vcObjects is None:
                    config['singleVisual']['objects'] = {}
                elif objects is None and vcObjects is not None:
                    config['singleVisual']['objects'] = config['singleVisual'].get('vcObjects')
                elif objects is not None and vcObjects is not None:
                    config['singleVisual']['objects'].update(config['singleVisual'].get('vcObjects'))

                visual_objects = self._fetch_object_properties_value(config['singleVisual']['objects'])

                page_visuals_properties['visuals'].append({
                    'visual_type': config['singleVisual']['visualType'],
                    'objects': visual_objects,
                })

            page_properties = self._get_page_properties(report_page_section)
            if page_properties is not None:
                page_objects = self._fetch_object_properties_value(page_properties.get('objects', {}))
                if page_objects is not None:
                    page_visuals_properties['visuals'].append({
                        'visual_type': 'page',
                        'objects': page_objects,
                    })
            return page_visuals_properties
        except Exception as e:
            ShowErrorDialog(LogException(e))

    def _fetch_object_properties_value(self, objects):
        try:
            visual_objects = {}
            for objectName, object_properties in objects.items():
                visual_objects[objectName] = {}
                for object_property_name, object_property_value in object_properties[0]['properties'].items():
                    if type(object_property_value) is dict:
                        if any(key in object_property_value for key in ['expr', 'solid']):
                            property_value = ''
                            if object_property_value.get('expr') is not None:
                                try:
                                    property_value = object_property_value.get('expr').get('Literal').get('Value')
                                    property_value = self._fixPropertyValue(property_value)
                                except AttributeError as e:
                                    # print(object_property_name, object_property_value)
                                    pass
                            elif object_property_value.get('solid') is not None:
                                try:

                                    color_value = object_property_value.get('solid').get('color').get('expr')
                                    if color_value.get('Literal') is not None:
                                        property_value = object_property_value.get('solid').get('color').get(
                                            'expr').get(
                                            'Literal').get(
                                            'Value')
                                        property_value = self._fixPropertyValue(property_value)
                                    elif color_value.get('ThemeDataColor') is not None:
                                        color = self._default_theme_accent_colors[color_value.get('ThemeDataColor')
                                            .get('ColorId')]
                                        shade_percent = color_value.get('ThemeDataColor').get('Percent')
                                        if color[0] is not '#':
                                            color = '#' + color
                                        ColorUtil.ShadeColor(color, shade_percent)
                                        property_value = ColorUtil.ShadeColor(color, shade_percent)
                                    property_value = {
                                        'solid': {
                                            'color': property_value
                                        }
                                    }
                                except AttributeError as e:
                                    # print(object_property_name, object_property_value)
                                    pass
                            visual_objects[objectName][object_property_name] = property_value
                    else:
                        visual_objects[objectName][object_property_name] = object_property_value[0]

            return visual_objects
        except Exception as e:
            ShowErrorDialog(LogException(e))

    def _get_page_properties(self, report_page_section):
        page_properties = json.loads(report_page_section['config'])
        return page_properties

    def _fixPropertyValue(self, value):
        try:
            if type(value) is str:
                if "'" in value:
                    value = value.strip("'")
                    value = value.replace("''", "")  # For font family

                if value[-1] in ['L', 'D'] and value[0] != '#':
                    value = int(float(value[0:-1]))
                    return value

                if value.isdigit():
                    value = int(float(value))
                    return value

                if value in ['true', 'false']:
                    if value == 'true':
                        value = True
                    else:
                        value = False
                    return value
            return value
        except Exception as e:
            ShowErrorDialog(LogException(e))

    def modifiedDataStructure(self):
        try:
            _reportVisualData = {}
            for reportSection in self.get_report_sections():
                _reportVisualData[reportSection['name']] = self._get_page_wise_visual_properties(reportSection)

            for reportPage in _reportVisualData:
                for visual in _reportVisualData[reportPage].get('visuals'):
                    visual['__selected'] = False
                    visualObjects = visual.get('objects')
                    for object in visualObjects:
                        visualObjects[object]['__selected'] = True

            return _reportVisualData
        except Exception as e:
            ShowErrorDialog(LogException(e))


if __name__ == '__main__':
    themeGenerator = PowerBIThemeGenerator('C:/test.pbix')
    theme_data = themeGenerator._get_report_layout_data()
    print(theme_data)
