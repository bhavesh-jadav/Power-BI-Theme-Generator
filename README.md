# Power BI Theme Generator

Power BI Theme Generator is a desktop tool which helps you generate themes from Power BI file. It scans the Power BI file for visuals 
and it extracts all available properties of visuals such as colors, formatting etc. and shows it to user. Then user selects which visuals and 
properties they want to extract from Power BI file and Power BI Theme Generator will generate `.json ` theme file which can be imported in another Power BI 
files to get same look and feel.

Standard approach for using Power BI Theme Generator is as below:
* Create on template Power BI file with visuals that you are going to use in others reports.
* Assign colors, format etc. to visuals in template Power BI file.
* Use Power BI Theme Generator tool on this template Power BI file and generate theme.
* Use this theme across other reports.

It is not necessary to create template Power BI file. You can use any Power BI file to extract visual properties to create them using Power BI Theme Generator.

## Getting Started

1. Clone this repository.
2. Install all libraries mentioned in `requirement.txt`
3. Use following ways to run the program
    * Run `PowerBIThemeGeneratorGUI.py` from `src/main/python`
    * From project root folder type `fbs run` in cmd to run the program

## Deployment

Run `fbs installer` to create installer file for program

## Built With

* [PyQt](https://riverbankcomputing.com/software/pyqt/intro) - Python wrapper around GUI framework Qt
* [fbs](https://build-system.fman.io/) - For creating final `.exe`
* [QT Designer](http://doc.qt.io/qt-5/qtdesigner-manual.html) - For designing UI

## Authors

* **Bhavesh Jadav** - *Initial work* - [bhavesh-jadav](https://github.com/bhavesh-jadav)

See also the list of [contributors](https://github.com/bhavesh-jadav/Power-BI-Theme-Generator/graphs/contributors) who participated in this project.

## License

This project is licensed under the GNU GENERAL PUBLIC LICENSE - see the [LICENSE](LICENSE) file for details

## Acknowledgments

* [stackoverflow](http://stackoverflow.com)
