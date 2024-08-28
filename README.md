# 🌟 Universe Reports Extractor

[![Java](https://img.shields.io/badge/Java-ED8B00?style=for-the-badge&logo=java&logoColor=white)](https://www.java.com/) [![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

## ✨ Overview

**SAP BO Universe Reports Extractor** is a Java utility designed to extract and list reports associated with a specific universe in SAP BusinessObjects. The program retrieves report details like Report Name, ID, Type, Owner, Universe Name, and Report Path, saving them in an Excel file.

## 🚀 Features

- Fetches reports linked to a specified universe.
- Exports report details to an Excel file.
- Simple configuration and easy to use.

## 🛠️ How to Use

### 1. Prerequisites

- Java Development Kit (JDK) 8 or higher.
- Apache POI (for Excel file manipulation).
- SAP BusinessObjects SDK libraries.

### 2. Clone the Repository

```bash
git clone https://github.com/swapnilyavalkar/SAP-BO-Universe-Reports-Extractor.git
cd SAP-BO-Universe-Reports-Extractor
```

### 3. Set Up Classpath for Libraries

Ensure the required libraries are included in your classpath:

```bash
javac -cp ".:/path/to/libs/*" UniverseReports.java
```

Replace `/path/to/libs` with the actual path where your SAP BusinessObjects SDK libraries and other dependencies are located.

### 4. Run the Program

Execute the following command, replacing the placeholders with actual values:

```bash
java -cp ".:/path/to/libs/*" UniverseReports <CMS_HOST> <USERNAME> <PASSWORD> <UNIVERSE_SI_ID>
```

Example:

```bash
java -cp ".:/path/to/libs/*" UniverseReports myCMSHost administrator myPassword 12345
```

The output will be saved as an Excel file in the directory specified in the code.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 👤 Author

Authored by [Swapnil Yavalkar](https://github.com/swapnilyavalkar).

## 🌟 Contributing

Contributions are welcome! Feel free to submit a pull request or open an issue.

---

Happy coding! 😊