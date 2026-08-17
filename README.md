# OLT MOP Automation Tool

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://your-app-url.streamlit.app)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## 📋 Overview

A comprehensive web-based tool for automating OLT (Optical Line Terminal) MOP (Method of Procedure) integration. It compares FIO (Excel), EWP (Image), and MOP (Word) files between two sites to identify differences and generate detailed reports.

## ✨ Features

- **📊 FIO Processing**: Parse Excel files containing OLT connection data
- **🖼️ EWP OCR**: Extract text from network diagrams and images using Tesseract OCR
- **📝 MOP Parsing**: Process Word documents with OLT configuration details
- **🔍 Difference Analysis**: Automatically compare data between two sites
- **📄 Report Generation**: Create Excel and PDF reports with detailed differences
- **🎯 Severity Classification**: Categorize differences as Critical, Warning, or Info
- **🔄 Cross-Reference Validation**: Verify consistency across different file types

## 🚀 Quick Start

### Local Development

1. **Clone the repository**
```bash
git clone https://github.com/yourusername/olt-mop-automation.git
cd olt-mop-automation
