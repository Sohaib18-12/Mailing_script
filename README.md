# Job Application Email Automation

Welcome! This code is designed to automate the process of sending job application emails using the `smtplib` module. Below are the steps to get started:

## Getting Started

1. **Replace Resume and Cover Letter**:
   - Replace the files `cover.pdf` and `Resume.pdf` with your own resume and cover letter.

2. **Fill Personal Information**:
   - Open the Excel file `Job_Track.xlsx` and navigate to the "Personal Data" sheet. Fill in your personal information.

3. **Get Gmail Key**:
   - Follow [this link](https://support.google.com/mail/answer/185833?hl=en) to obtain your Gmail API key.

4. **Set Gmail Password**:
   - Copy and paste the generated password into the "Password" cell in the Excel file under the "Personal Data" sheet.

5. **Create Company Database**:
   - Start creating your company database in the "data" sheet of the Excel file.

6. **Customize Email Message**:
   - The file `Body.py` defines the message that will be sent. Customize it according to your preferences. Knowledge of HTML can be helpful for advanced customization.

7. **Send Emails**:
   - Open the file `main.py` and run it to start sending the emails. The emails will be sent one by one.

## Important Notes

- **Test Emails First**:
  - It's recommended to test with emails you have access to before sending real job applications. This ensures everything is working correctly.

## Feedback and Support

If you have any questions, suggestions for improvements, or need assistance, feel free to contact me. I'm constantly working on enhancing this tool to make it more user-friendly for everyone.

Take care and happy job hunting! :)
