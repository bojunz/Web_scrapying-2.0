#!/usr/bin/env python
# coding: utf-8

# In[1]:


import sys
import json
import time
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLineEdit,QComboBox, QTextEdit, QLabel, QVBoxLayout, QHBoxLayout, QFormLayout, QWidget, QGroupBox, QMessageBox
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPalette, QColor,QIcon
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import threading
from PyQt5.QtWidgets import QShortcut
from PyQt5.QtGui import QKeySequence
from datetime import datetime
from PyQt5.QtWidgets import QSpacerItem, QSizePolicy

class AmazonAutomationApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Amazon Automation Tool")
        self.setGeometry(100, 100, 1000, 600)  # Adjusted width for a wider layout
        self.setFixedSize(self.size())
        self.results = []
        #set up logo
        self.setWindowTitle('Timetec')
        self.setWindowIcon(QIcon("tele.ico"))
        self.show()


        # Set up background color
        palette = self.palette()
        palette.setColor(QPalette.Window, QColor("#f0f2f5"))  # Light background
        self.setPalette(palette)

        # Main layout
        main_layout = QHBoxLayout()

        # Left Panel - Inputs and Controls
        left_layout = QVBoxLayout()

        # Login Section with GroupBox for structure
        login_box = QGroupBox("Login Information")
        login_layout = QFormLayout()
        
        # Email and Password Fields
        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText("Enter your email")
        self.email_input.setFixedWidth(300)
        
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setPlaceholderText("Enter your password")
        self.password_input.setFixedWidth(300)
        
        # Tip Label
        self.tip_label = QLabel("Please enter your username and password before you start the app.")
        self.tip_label.setStyleSheet("font-size: 10px; color: #555;")
        login_layout.addWidget(self.tip_label)

        # Don't Know Your Account Button
        self.forgot_button = QPushButton("Don't know your account?")
        self.forgot_button.setStyleSheet("color: blue; text-decoration: underline;")
        self.forgot_button.setFlat(True)
        self.forgot_button.clicked.connect(self.show_account_info)

        # Add email, password, and forgot button to login layout
        login_layout.addRow(QLabel("Email:"), self.email_input)
        login_layout.addRow(QLabel("Password:"), self.password_input)
        
        #country
        self.country_dropdown = QComboBox()
        self.country_dropdown.addItem("Please select country")  # Add the prompt first
        self.country_dropdown.addItems(["United States", "Canada", "France", "Italy", 'Germany', 'Spain' , 'United Kingdom'])
        self.country_dropdown.setFixedWidth(300)
        # Optionally disable the first item so it can’t be selected again
        self.country_dropdown.model().item(0).setEnabled(False)
        login_layout.addRow(QLabel("Country:"), self.country_dropdown)
        
        def print_selected_country():   
            selected_country = self.country_dropdown.currentText()
            print(f"Selected Country: {selected_country}")
            
        self.country_dropdown.currentIndexChanged.connect(print_selected_country)



        # Create a horizontal layout for the button
        forgot_button_layout = QHBoxLayout()
        forgot_button_layout.addWidget(self.forgot_button)
        forgot_button_layout.setAlignment(Qt.AlignRight)
        login_layout.addRow(forgot_button_layout)

        # Add the layout to the form layout
        login_layout.addRow(forgot_button_layout)
        
        login_box.setLayout(login_layout)
        left_layout.addWidget(login_box)

        # Order ID input with label
        self.order_id_label = QLabel("Order IDs (comma-separated):")
        left_layout.addWidget(self.order_id_label)
        self.order_id_input = QLineEdit()
        self.order_id_input.setPlaceholderText("e.g., 114-3523692-8769050, 113-8092055-4566625")
        self.order_id_input.setFixedHeight(50)
        left_layout.addWidget(self.order_id_input)
        


        # Custom Message Input
        self.custom_message_label = QLabel("Custom Message to Customer:")
        left_layout.addWidget(self.custom_message_label)
        self.custom_message_input = QTextEdit()
        self.custom_message_input.setPlaceholderText("Enter your custom message here...")
        self.custom_message_input.setFixedHeight(200)
        left_layout.addWidget(self.custom_message_input)

        # Technical Support Area
        support_box = QGroupBox("Technical Support")
        support_layout = QVBoxLayout()
        self.support_label = QLabel("For any questions or help, contact: bojunz@timetecinc.com")
        self.support_label.setStyleSheet("font-style: italic; color: #555;")
        support_layout.addWidget(self.support_label)
        support_box.setLayout(support_layout)
        left_layout.addWidget(support_box)

        # Start Button
        self.start_button = QPushButton("Start Automation")
        self.start_button.clicked.connect(self.start_automation)
        self.start_button.setFixedWidth(200)
        self.start_button.setStyleSheet("background-color: #4CAF50; color: white; padding: 5px; font-weight: bold;")
        left_layout.addWidget(self.start_button, alignment=Qt.AlignCenter)
        
        
        # Bind Enter key (Return key)
        enter_shortcut = QShortcut(QKeySequence(Qt.Key_Return), self)
        enter_shortcut.activated.connect(self.start_button.click)

        # Optional: also bind keypad Enter
        numpad_enter_shortcut = QShortcut(QKeySequence(Qt.Key_Enter), self)
        numpad_enter_shortcut.activated.connect(self.start_button.click)

        # Right Panel - Logs
        right_layout = QVBoxLayout()

        # Order Information Output
        order_info_box = QGroupBox("Order Information Log")
        order_info_layout = QVBoxLayout()
        self.order_output = QTextEdit()
        self.order_output.setReadOnly(True)
        self.order_output.setFixedHeight(150)  # Adjusted height
        self.order_output.setStyleSheet("background-color: #e8eaf6; color: #333; font-family: Consolas;")
        order_info_layout.addWidget(self.order_output)
                
            
        # Create the buttons
        clear_button = QPushButton("Clear Content")
        clear_button.setFixedWidth(150)
        clear_button.setStyleSheet("background-color: #99FFCC; color: blue; ")
        clear_button.clicked.connect(self.clear_order_output)

        export_button = QPushButton("Export to CSV")
        export_button.setFixedWidth(150)
        export_button.setStyleSheet("background-color: #FFCC99; color: black; ")
        export_button.clicked.connect(self.export_to_csv)

        # Horizontal layout for buttons with spacer in between
        button_layout = QHBoxLayout()
        button_layout.addWidget(export_button)

        # Add spacer to push the next button to the right
        button_layout.addItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))

        button_layout.addWidget(clear_button)

        # Add the button layout to the bottom of the order info layout
        order_info_layout.addLayout(button_layout)

        # Set final layout
        order_info_box.setLayout(order_info_layout)
        right_layout.addWidget(order_info_box)            
            

        # Logging Output
        log_box = QGroupBox("General Log")
        log_layout = QVBoxLayout()
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setFixedHeight(300)  # Increased height
        self.log_output.setStyleSheet("background-color: #e8eaf6; color: #333; font-family: Consolas;")
        log_layout.addWidget(self.log_output)
        log_box.setLayout(log_layout)
        right_layout.addWidget(log_box)

        # Add left and right layouts to the main layout
        main_layout.addLayout(left_layout)
        main_layout.addLayout(right_layout)

        # Set layout in the central widget
        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)
    
    def export_to_csv(self):
        import csv
        from PyQt5.QtWidgets import QFileDialog

        if not self.results:
            QMessageBox.warning(self, "No Data", "There are no results to export.")
            return

        # Prompt user for save location
        file_path, _ = QFileDialog.getSaveFileName(self, "Save CSV", "", "CSV Files (*.csv)")

        if file_path:
            try:
                with open(file_path, mode='a+', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    writer.writerow(["Order ID", "Status", "Date"])  # Header
                    writer.writerows(self.results)
                QMessageBox.information(self, "Export Complete", f"Results exported successfully to:\n{file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Export Failed", f"Failed to export CSV:\n{e}")
        
    def clear_order_output(self):
        self.order_output.clear()
        self.log("Order Information Output cleared.")

    def log(self, message):
        self.log_output.append(message)
        self.log_output.verticalScrollBar().setValue(self.log_output.verticalScrollBar().maximum())
    
    def log_order_info(self, message):
        self.order_output.append(message)
        self.order_output.verticalScrollBar().setValue(self.order_output.verticalScrollBar().maximum())
        
    def get_shadow_root(self, driver, element):
        return driver.execute_script('return arguments[0].shadowRoot', element)

    def show_account_info(self):
        # Display a pop-up message with account information
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setWindowTitle("Account Information")
        msg_box.setText("Please use your Timetec account to login.")
        msg_box.exec_()
    
    def start_automation_thread(self):
        # Run start_automation in a separate thread
        self.thread = threading.Thread(target=self.start_automation)
        self.thread.start()
    
    
    def custom_message(self,order_id):
        self.SKU = self.driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr/td[3]/div/div[1]/div/a/div').text
                
        self.custom_message_en = f'''Hi there,

Thank you for shopping with Timetec through Amazon. We welcome you to our Timetec family as a valued customer! 

As part of the Timetec Customer Satisfaction Team, I wanted to reach out to learn whether you have received the item and how it has been working for you.

Here are the details for this Amazon order:{order_id} 
{self.SKU}

Meanwhile, we as a company would be deeply grateful for your unbiased feedback for us on Amazon.

We would sincerely appreciate it if you could take the time to share your model or manufacture system part No. so that others can benefit from your experience.

If there’s anything else I can do for you, please feel free to let me know.

Sincerely,
Timetec International Inc.
J.Z
Customer Service Team '''
                
        self.custom_message_es = f'''Hola,

Gracias por comprar en Timetec a través de Amazon. ¡Te damos la bienvenida a nuestra familia Timetec como un cliente valioso!

Como parte del Equipo de satisfacción del cliente de Timetec, quería comunicarme contigo para saber si has recibido el artículo y cómo te ha funcionado.

A continuación, se muestran los detalles de este pedido de Amazon:{order_id} - {self.SKU}

Mientras tanto, como empresa estaríamos profundamente agradecidos por tus comentarios imparciales para nosotros en Amazon.

Apreciaríamos sinceramente que te tomaras el tiempo de compartir tu modelo o el número de pieza del sistema de fabricación para que otros puedan beneficiarse de tu experiencia.

Si hay algo más que pueda hacer por ti, no dudes en hacérmelo saber.

Atentamente,
Timetec International Inc.
J.Z
Equipo de atención al cliente'''
    
    
    
    def start_automation(self):
        order_ids = self.order_id_input.text().split()
        email = self.email_input.text()
        password = self.password_input.text()
        P_custom_message = self.custom_message_input.toPlainText()

        #if not order_ids or not email or not password or not custom_message:
            #self.log("Please provide Order IDs, Email, Password, and a Custom Message.")
            #return
        
        self.log("Program started...")

        # Set up ChromeDriver and Selenium options
        options = webdriver.ChromeOptions()
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36')
        
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

        # Login process
        try:
            s_url = 'https://sellercentral.amazon.com/orders-v3/order/114-3523692-8769050'
            self.driver.get(s_url)

            # Add cookies
            cookie_dict = {
                'domain': '.amazon.com',
                'expiry': 1758228517,
                'httpOnly': True,
                'name': 'at-main',
                'path': '/',
                'sameSite': 'Lax',
                'secure': True,
                'value': 'YourCookieValue'
            }

            self.driver.add_cookie(cookie_dict)
            self.log("Cookies loaded and page refreshed.")

            # Perform login
            WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.ID, 'ap_email')))
            self.driver.find_element(By.ID, 'ap_email').send_keys(email)
            self.driver.find_element(By.ID, 'continue').click()
            self.driver.find_element(By.ID, 'ap_password').send_keys(password)
            self.driver.find_element(By.ID, 'signInSubmit').click()
            self.log("Logged in successfully.")
            
            
            time.sleep(2)
            
            
            # Wait for the company selection
            selected_country = self.country_dropdown.currentText()
            print(selected_country)
            if selected_country in ['United States', 'Canada']:       
                WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.XPATH,"//*[text()='Timetec International Inc']")))
                self.driver.find_element(By.XPATH, "//*[text()='Timetec International Inc']").click()

                # Select country and confirm
                time.sleep(2)
                self.driver.find_element(By.XPATH, f"//*[text()='{selected_country}']").click()
            elif selected_country in ['France', 'Germany', 'Spain', 'United Kingdom', 'Italy']:
                WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.XPATH,"//*[text()='Timetec Inc Europe']")))
                self.driver.find_element(By.XPATH, "//*[text()='Timetec Inc Europe']").click()

                # Select country and confirm
                time.sleep(2)
                self.driver.find_element(By.XPATH, f"//*[text()='{selected_country}']").click()
            #confirm
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//*[contains(text(),'Select account')]").click()



            # Process each order ID
            for order_id in order_ids:
                #order_url = f'https://sellercentral.amazon.com/orders-v3/order/{order_id.strip()}'
                order_url = 'https://sellercentral.amazon.com/orders-v3/order/{}'.format(order_id.strip("'"))
                time.sleep(2)   
                
                               
                
                try:
                    self.driver.get(order_url)
                    self.log(f"Processing order: {order_id.strip()}")
                    
                    #load custom message
                    self.custom_message(order_id)

                    # Click on the buyer
                    WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div/div/div[1]/div[1]/div/div[2]/div[2]/div/div/div/div/div[2]/div/table/tbody/tr/td[2]/span/b/a')))
                    self.driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div/div/div[1]/div[1]/div/div[2]/div[2]/div/div/div/div/div[2]/div/table/tbody/tr/td[2]/span/b/a').click()
                                
                    # Switch to new window
                    self.driver.implicitly_wait(20)
                    time.sleep(3)
                    self.driver.switch_to.window(self.driver.window_handles[-1])
                    


                    # Interact with shadow DOM, this shoud be the send feedback button(other)
                    refund_button = self.driver.find_element(By.XPATH, '(//*[@id="ayb-contact-buyer"]/div[3]/kat-box/div/kat-radiobutton)[last()]')
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", refund_button)
                    self.driver.execute_script("arguments[0].click();", refund_button)
                    self.log("Clicked refund button.")
                    
                    
                    #check the customer language
                    time.sleep(1)
                    languages = self.driver.find_element(By.XPATH, '//*[@id="ayb-contact-buyer"]/div[4]/kat-tooltip/kat-label').text
                    print(languages)
                    

                    #choose corresponding languages
                    if P_custom_message !='':
                        custom_message = P_custom_message
                    else:
                        if languages == "Buyer's Language of Preference: Spanish":
                            print('spanish')
                            custom_message = self.custom_message_es
                        else:
                            print('english')
                            custom_message = self.custom_message_en
                

                    # Locate shadow DOM element for textarea
                    #Need to use JS to enter characters directly in the code
                    shadow_host = self.driver.find_element(By.XPATH, '//*[@id="ayb-contact-buyer"]/div[4]/form/div[1]/div/div[1]/kat-textarea')
                    shadow_root = self.driver.execute_script('return arguments[0].shadowRoot', shadow_host)
                    #custom_message = str(custom_message)
                    #self.driver.execute_script("arguments[0].setAttribute('value', '{}');".format(custom_message), shadow_host)
                    
                    #Need to use JS to enter characters directly in the code
                    self.driver.execute_script("arguments[0].setAttribute('value', arguments[1]);", shadow_host, custom_message)

                    self.log("Message set in the textarea.")
                    self.log(f"Message send successful,{order_id}")

                    # Dispatch input event for validation
                    time.sleep(2)
                    self.driver.execute_script("""
                        let event = new Event('input', { bubbles: true });
                        arguments[0].dispatchEvent(event);
                    """, shadow_host)

                    # Click the send button using CSS selector
                    time.sleep(3)
                    shadow_host = self.driver.find_element(By.XPATH, '//*[@id="ayb-contact-buyer"]/div[8]/kat-button')
                    shadow_root = self.get_shadow_root(self.driver, shadow_host)
                    submit_button = shadow_root.find_element(By.CSS_SELECTOR, 'button')
                    self.driver.execute_script("arguments[0].click();", submit_button)
                    
                    # Some customers use different language, need to confirm again 
                    time.sleep(2)
                    
                    #My attempt
                    #Combining shadow and wait
                    #Judge whether there is an additional click window, if so, click it additionally
                    if languages == "Buyer's Language of Preference: Spanish":
                        shadow_host_lan = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, '//*[@id="ayb-contact-buyer"]/div[6]/kat-modal/div/kat-button'))
                        )
                        print(shadow_host_lan)
                        print(bool(shadow_host_lan))
                        #shadow_host_lan = self.driver.find_element(By.XPATH, '//*[@id="ayb-contact-buyer"]/div[6]/kat-modal/div/kat-button')
                        shadow_root_lan = self.get_shadow_root(self.driver, shadow_host_lan)


                        submit_button = shadow_root_lan.find_element(By.CSS_SELECTOR, 'button')
                        self.driver.execute_script("arguments[0].click();", submit_button)

                         
                    self.log_order_info(f'Message sent successfully for order {order_id.strip()}')
                    self.results.append([order_id.strip(), 'Success', datetime.now().strftime('%Y-%m-%d')])

                except Exception as e:
                    self.log(f"An error occurred while interacting with shadow DOM: {e}")
                    self.log_order_info(f'Failed to send message for order {order_id.strip()}')
                    self.results.append([order_id.strip(), 'Failed', datetime.now().strftime('%Y-%m-%d')])
                time.sleep(3)

            self.log("All orders processed.")
        except Exception as e:
            self.log(f"An error occurred during the login or processing: {e}")
        finally:
            self.driver.quit()

# Run the application
app = QApplication(sys.argv)
window = AmazonAutomationApp()
window.show()
sys.exit(app.exec_())


# In[ ]:




