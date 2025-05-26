# Sea Level Report V2
[![GitHub Actions Status](https://github.com/linz/template-python-hello-world/workflows/Build/badge.svg)](https://github.com/linz/template-python-hello-world/actions)
[![Coverage: 100% branches](https://img.shields.io/badge/Coverage-100%25%20branches-brightgreen.svg)](https://pytest.org/)
[![Kodiak](https://badgen.net/badge/Kodiak/enabled?labelColor=2e3a44&color=F39938)](https://kodiakhq.com/)
[![Dependabot Status](https://badgen.net/badge/Dependabot/enabled?labelColor=2e3a44&color=blue)](https://github.com/linz/template-python-hello-world/network/updates)
[![License](https://badgen.net/github/license/linz/template-python-hello-world?labelColor=2e3a44&label=License)](https://github.com/linz/template-python-hello-world/blob/master/LICENSE)
[![Conventional Commits](https://badgen.net/badge/Commits/conventional?labelColor=2e3a44&color=EC5772)](https://conventionalcommits.org)
[![Code Style](https://badgen.net/badge/Code%20Style/black?labelColor=2e3a44&color=000000)](https://github.com/psf/black)
[![Imports: isort](https://img.shields.io/badge/%20imports-isort-%231674b1?style=flat&labelColor=ef8336)](https://pycqa.github.io/isort/)
[![Checked with mypy](http://www.mypy-lang.org/static/mypy_badge.svg)](http://mypy-lang.org/)
[![Code Style: prettier](https://img.shields.io/badge/code_style-prettier-ff69b4.svg)](https://github.com/prettier/prettier)
## Create PDF and Doc reports from the SLIM output.
### V1. Convert VB Script to JS with Macro -- Issues with conversion performance and support from external vendor(IIC)
```
Manual conversion with MS word XLS and VB
Old VB script has performance and system memory issue
Fully automated with Macro tool 
```
### V2. Convert JS to Python -- Issues with MS printer driver
```
Some Windows 10 users have reported printing issues after installing certain updates,
including duplicate copies, applications failing to print, and interrupted printing operations,
particularly after updates like KB5015807 and KB5014666.

Independent of the pdf printer driver
Generate Word files first and convert PDFs from the Word files 
```

### Getting started

1. Copy executable files from the Hydro repository (N:\Software\HPD\SeaLevelReport)
    - SeaLevelReport.exe
    - config.yaml
2. Edit the configuration file (config.yaml)
    ```
    folder_path: 'C:\\CSV files\\'                        --  Tide csv file location
    output_folder: 'C:\\Reports\\'                        --  Word and PDF out loaction
    linz_logo_path: 'C:\\linz_colour_cmyk_66mm_png.png'   --  LINZ logo file loaction
    ```
3. Open windows CMD

4. Change directory into the project directory (exe file saved)

5. execute the SeaLevelReport.exe