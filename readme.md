# Certificate Generator Application

## Introduction
This Certificate Generator Application allows users to generate certificates from a CSV file and create downloadable PDFs. Additionally, users can email these certificates using the macros provided in the `Mail.xlsm` file.

## Features
- Generate certificates from a CSV file.
- Downloadable PDF certificates.
- Option to email certificates using the provided Excel macro.

## Getting Started

### Prerequisites
- Python 3.x
- pip (Python package installer)

### Installation
1. **Clone the Repository**  
   Clone the repository to your local machine using:
   ```bash
   git clone https://github.com/anonymousknight07/Certificate-generator.git
   
2. Install Requirements
   Navigate to the cloned repository directory and install all the required packages using:
   ```bash
   pip install -r requirements.txt

### Running 
1. Start the Application
  Open the terminal and run the following command:
 ```bash
python wsgi.py
```

2. Access the Application
   Open your web browser and go to the respective port number (typically http://127.0.0.1:5000)

### Output
- The generated certificates will be saved on your local machine.
- To email the certificates, run the macros in the Mail.xlsm file located in the Data folder.
  
### Usage
1. Place your CSV file with the required details (like names, email addresses, etc.) in the appropriate location.
2. Run the application as described above.
3. Download the generated PDF certificates.
4. (Optional) To send the certificates via email, open the Mail.xlsm file in the Data folder and run the macros.

### Contributing
If you would like to contribute to this project, please fork the repository and submit a pull request.

### License
This project is licensed under the MIT License. See the LICENSE file for more details.

### Snapshot
![image](https://github.com/user-attachments/assets/960f0963-0373-44f3-970d-150159dfa46b)





