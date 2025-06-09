# # This code will remove the background noise and return the image of main captcha text (Updated Version).
# import cv2
# import numpy as np
# import pytesseract
# import matplotlib.pyplot as plt
# import re

# # Optional: set path to Tesseract manually if needed
# # pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# def preprocess_image(image_path):
#     """Preprocess the CAPTCHA image for better character recognition."""
#     try:
#         img = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
#         if img is None:
#             print("❌ Error: Image not found!")
#             return None

#         # Apply Gaussian blur to reduce noise
#         img = cv2.GaussianBlur(img, (5, 5), 0)

#         # Apply adaptive thresholding
#         _, img = cv2.threshold(img, 180, 255, cv2.THRESH_BINARY)

#         # Dilation to thicken characters
#         kernel = np.ones((2, 2), np.uint8)
#         img = cv2.dilate(img, kernel, iterations=1)

#         # Resize to a standard size
#         img = cv2.resize(img, (200, 50))

#         # Debugging: Show the processed image
#         plt.imshow(img, cmap='gray')
#         plt.title("Processed CAPTCHA Image")
#         plt.axis('off')
#         plt.show()

#         return img
#     except Exception as e:
#         print(f"❌ Error in preprocessing: {e}")
#         return None

# def extract_text_from_image(img):
#     """Extract exactly 6 alphanumeric characters from the processed image."""
#     try:
#         # Extract text with whitelist of letters and digits
#         config = '--psm 8 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'
#         raw_text = pytesseract.image_to_string(img, config=config)
        
#         # Clean the text: Remove unwanted characters
#         cleaned_text = re.sub(r'[^A-Za-z0-9]', '', raw_text)
        
#         # Take exactly 6 characters
#         if len(cleaned_text) >= 6:
#             final_text = cleaned_text[:6]
#         else:
#             # If less than 6 characters, you can either pad or just return whatever found
#             final_text = cleaned_text.ljust(6, '')  # Example: pad with underscores ""
        
#         return final_text
#     except Exception as e:
#         print(f"❌ Error in text extraction: {e}")
#         return None

# if __name__ == "__main__":
#     image_path = "captcha.png"  # Input image path
#     processed_img = preprocess_image(image_path)
    
#     if processed_img is not None:
#         extracted_text = extract_text_from_image(processed_img)
        
#         if extracted_text:
#             print(f"Final Extracted Text (6 chars): {extracted_text}")
#         else:
#             print("❌ Text extraction failed.")
        
#         output_path = "processed_output.png"
#         cv2.imwrite(output_path, processed_img)
#         print(f"Preprocessing complete. Processed image saved as '{output_path}'.")
#         print(extracted_text)
#     else:
#         print("❌ Preprocessing failed.")
# test.py

import cv2
import numpy as np
import matplotlib.pyplot as plt

def preprocess_image(image_path, output_path="processed_captcha.png"):
    try:
        img = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
        if img is None:
            raise ValueError("Image not found")

        # Noise reduction
        img = cv2.GaussianBlur(img, (5, 5), 0)
        _, img = cv2.threshold(img, 180, 255, cv2.THRESH_BINARY)
        kernel = np.ones((2, 2), np.uint8)
        img = cv2.dilate(img, kernel, iterations=1)
        img = cv2.resize(img, (200, 50))  # Resize to consistent input

        # Save processed image
        cv2.imwrite(output_path, img)
        plt.imshow(img, cmap='gray')
        plt.title("Processed CAPTCHA Image")
        plt.axis('off')
        plt.show()
        return output_path
    except Exception as e:
        print(f"❌ Error in preprocessing: {e}")
        return None

if __name__ == "__main__":
    preprocess_image("captcha.png")

