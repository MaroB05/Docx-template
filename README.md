# Docx Template

Fill in your Word document templates with a sleek GUI - because manually replacing placeholders is so 2005.

## Overview

Docx Template is a simple yet powerful tool that allows you to quickly fill out docx templates through a friendly user interface. No more manually hunting for placeholders - this app finds them all and presents them as a neat form for you to fill out.

Think of it as a document secretary that doesn't need coffee breaks or complain about papercuts.

## Features

- **Template Selection**: Easily choose from your available templates
- **Dynamic Form Generation**: Automatically identifies all placeholders in your template
- **Custom Output Naming**: Save your filled documents with custom filenames
- **Simple Interface**: No complex menus or confusing options

## How It Works

1. Place your template docx files in the `data/` folder
2. Run the application
3. Select your template from the dropdown
4. Fill in the fields that appear
5. Name your output file
6. Click submit
7. Find your completed document in the `output/` folder

## Requirements

- Python 3.6+
- PyQt6
- Spire.Doc

## Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/Docx-template.git
cd Docx-template

# Install dependencies
pip install PyQt6 spire.doc
```

## Usage

```bash
python main.py
```

## Template Format

The application looks for placeholders in your Word documents that follow this format:
```
*PlaceholderName*
```

For example, if your document contains `*Name*`, the application will create a field labeled "Name" for you to fill in.

## Directory Structure

```
Docx-template/
├── data/           # Place your .docx templates here
├── output/         # Filled documents will be saved here
├── main.py         # The main application script
└── README.md       # This helpful file you're reading now
```

## Contributing

Found a bug? Have a feature idea? Pull requests are welcome! Just make sure your code doesn't have as many placeholders as the templates it's filling.

## License

MIT - Use it, improve it, share it. Just don't blame me when your boss keeps asking you to make "just one more quick template."

---

*Happy templating!*
