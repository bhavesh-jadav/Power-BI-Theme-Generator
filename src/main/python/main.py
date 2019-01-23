from fbs_runtime.application_context import ApplicationContext
from PowerBIThemeGeneratorGUI import PowerBIThemeGeneratorWindow

import sys


class AppContext(ApplicationContext):           # 1. Subclass ApplicationContext
    def run(self):                              # 2. Implement run()
        window = PowerBIThemeGeneratorWindow()
        return self.app.exec_()                 # 3. End run() with this line


if __name__ == '__main__':
    appctxt = AppContext()                      # 4. Instantiate the subclass
    exit_code = appctxt.run()                   # 5. Invoke run()
    sys.exit(exit_code)
