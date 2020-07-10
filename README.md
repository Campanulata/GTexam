# GTexam

此项目使用BHCexam<https://github.com/mathedu4all/bhcexam>提供的cls模板构建

py主程序读取docx后写入两个tex文件中，由main.tex进行编译

vscode设置备份:settings.json

{
  "workbench.colorTheme": "Monokai",
  "sync.autoDownload": true,
  "sync.autoUpload": true,
  "editor.tabSize": 4,
  "editor.renderControlCharacters": true,
  "editor.renderWhitespace": "all",
  "git.enableSmartCommit": true,
  "git.confirmSync": false,
  "markdownlint.config": {
    "MD029": false,
    "MD045": false
  },
  "explorer.confirmDragAndDrop": false,
  "editor.codeActionsOnSave": {
    "source.fixAll.markdownlint": true
  },
  "window.zoomLevel": 0,
  "explorer.confirmDelete": false,
  "liveServer.settings.donotShowInfoMsg": true,
  "python.dataScience.askForKernelRestart": false,
  "latex-workshop.latex.tools": [
    {
      "name": "latexmk",
      "command": "latexmk",
      "args": [
        "-synctex=1",
        "-interaction=nonstopmode",
        "-file-line-error",
        "-pdf",
        "%DOC%"
      ]
    },
    {
      "name": "xelatex",
      "command": "xelatex",
      "args": [
        "-synctex=1",
        "-interaction=nonstopmode",
        "-file-line-error",
        "%DOC%"
      ]
    },
    {
      "name": "pdflatex",
      "command": "pdflatex",
      "args": [
        "-synctex=1",
        "-interaction=nonstopmode",
        "-file-line-error",
        "%DOC%"
      ]
    },
    {
      "name": "bibtex",
      "command": "bibtex",
      "args": [
        "%DOCFILE%"
      ]
    }
  ],
  "latex-workshop.latex.recipes": [
    {
      "name": "xelatex",
      "tools": [
        "xelatex"
      ]
    },
    {
      "name": "latexmk",
      "tools": [
        "latexmk"
      ]
    },
    {
      "name": "pdflatex -> bibtex -> pdflatex*2",
      "tools": [
        "pdflatex",
        "bibtex",
        "pdflatex",
        "pdflatex"
      ]
    }
  ],
  "latex-workshop.latex.clean.fileTypes": [
    ".aux",
    ".bbl",
    ".blg",
    ".idx",
    ".ind",
    ".lof",
    ".lot",
    ".out",
    ".toc",
    ".acn",
    ".acr",
    ".alg",
    ".glg",
    ".glo",
    ".gls",
    ".ist",
    ".fls",
    ".log",
    "*.fdb_latexmk"
  ],
  "latex-workshop.view.pdf.viewer": "tab",
  "latex-workshop.view.pdf.external.viewer.command": "/Applications/Skim.app/Contents/SharedSupport/displayline",
  "latex-workshop.view.pdf.external.viewer.args": [
    "-r",
    "0",
    "%PDF%"
  ],
  "latex-workshop.view.pdf.external.synctex.command": "/Applications/Skim.app/Contents/SharedSupport/displayline",
  "latex-workshop.view.pdf.external.synctex.args": [
    "-r",
    "%LINE%",
    "%PDF%",
    "%TEX%"
  ],
  "editor.wordWrap": "on",
  "workbench.startupEditor": "newUntitledFile",
  "workbench.iconTheme": "Monokai Pro (Filter Machine) Icons",
  "latex-workshop.synctex.afterBuild.enabled": true,
  "gitlens.advanced.messages": {
    "suppressLineUncommittedWarning": true
  },
  "python.languageServer": "Pylance",
}
