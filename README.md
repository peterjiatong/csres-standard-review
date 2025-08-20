# Chinese Standards Review System (csres-standard-review)

[![Python](https://img.shields.io/badge/Python-3.12+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](https://claude.ai/chat/LICENSE)

A comprehensive toolkit for crawling, managing, and validating Chinese national standards from the official China Standards Research (csres.com) website. This project provides automated tools for maintaining a local standards database and checking compliance in technical reports.

## 🚀 Features

### 1. Standards Database Management

* **Crawling** : Fetch latest standard information from csres.com
* **Data Validation** : Verify standard codes, names, and status information, monitor superseded standards and their replacements

### 2. Report Compliance Checking

* **Document Parsing** : Extract standards from Word documents (.docx)
* **Validation Engine** : Check standard validity, naming consistency, and current status
* **Batch Processing** : Process multiple reports simultaneously
* **Detailed Reports** : Generate comprehensive compliance reports

Also with a Comprehensive Logging

## 🛠️ Installation

### Prerequisites

* Python 3.12+
* Git

### Setup Steps

1. **Clone the repository**
   ```bash
   git clone https://github.com/peterjiatong/csres-standard-review.git
   cd csres-standard-review
   ```
2. **Create virtual environment**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```
3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```
4. **Create required directories**
   ```bash
   mkdir reports
   ```

## ⚙️ Configuration

Create a `.env` file with the following variables:

```bash
# CSRes Website Credentials
CSRES_USERNAME=your_email@example.com
CSRES_PASSWORD=your_password_hash

# File Paths
SRC=standards.xlsx
DEST=standards.xlsx
```

note:

1. Recommand to buy a membership at csres.com
2. I won't be able to provide a standards.xlsx here, feel free to generate one for your own

## 🚦 Usage

### 1. Update Standards Database

```bash
python update_database_excel.py
```

**What it does:**

* Fetches latest information for all standards in the database
* Updates status, dates, and replacement information
* Generates update reports and logs
* Handles network errors and rate limiting

**Requirements:**

* Ensure there is a SRC and DEST xlsx exist in the project directory
* Close all Excel files before running
* Stable internet connection

**Output files:**

* `更新数据库.exe的运行结果_MM_DD_N/`
  * `标准更新报告.txt `- Summary of new standards added

**Log Files:**

* `log/` - Detailed execution logs
* `log_excel/` - Excel format logs for debugging

### 2. Check Standards in Reports

```bash
python check_standards_in_reports.py
```

**What it does:**

* Scans all `.docx` files in the `reports/` folder
* Extracts standard references using regex patterns
* Validates against the local database
* Generates compliance reports

**Requirements:**

* Place report files in `reports/` folder
* Ensure all Word documents are closed
* Run database update first for latest information

**Output files:**

* `检查报告中的标准.py的运行结果_MM_DD_N/`
  * `标准检查报告.txt` - Detailed compliance report
  * `标准更新报告.txt` - New standards found in reports

**Note:**

The system recognizes standards in these formats only:

* `GB/T 5750-2006 《生活饮用水标准检验方法》`
* `《建筑设计防火规范》 GB 50016-2014`
* `(GB 12801-2008) 《生产过程安全卫生要求总则》`

**Log Files:**

* `log/` - Detailed execution logs
* `log_excel/` - Excel format logs for debugging

## 🚧 Known Limitations

1. Intricate table layouts may cause parsing errors
2. Requires stable internet connection
3. Excel files must be closed during operation
4. The system won't be able to determine standard name without 《》

## 🔄 Version History

### v0.3 (Current)

* ✅ Added replacement information for superseded standards
* ✅ Improved database schema with replacement tracking
* ✅ Enhanced error handling and logging

### v0.2

* ✅ Fixed table content matching issues
* ✅ Added support for line breaks and soft returns
* ✅ Improved console output formatting

### v0.1

* ✅ Initial release with basic crawling functionality
* ✅ Report generation and folder organization
* ✅ Standard format validation

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](https://claude.ai/chat/LICENSE) file for details.

## 🏢 Acknowledgments

* **China Standards Research (CSRes)** for providing the standards database
* **Centre Testing International Group Corporation** for project sponsorship
* **BeautifulSoup** and **pandas** communities for excellent documentation

## 📞 Support

If you find this project helpful, please consider giving it a ⭐ on GitHub!

For issues and questions:

* 📧 Create an issue on GitHub
* 🔍 Review log files for debugging information

---

 **Note** : This tool is designed for legitimate research and compliance checking purposes. Please respect the terms of service of the CSRes website and use responsibly.
