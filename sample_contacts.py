#!/usr/bin/env python3
"""
Script to create a sample Excel file with the required format for the mail sender application.
Run this script to generate a sample_contacts.xlsx file.
"""

import pandas as pd

def create_sample_excel():
    """Create a sample Excel file with the required format"""
    # Sample data with the required columns
    data = {
        'genre': ['m', 'f', 'monsieur', 'madame', 'm'],
        'nom': ['Dupont', 'Martin', 'Bernard', 'Dubois', 'Moreau'],
        'email': [
            'jean.dupont@example.com',
            'marie.martin@example.com', 
            'pierre.bernard@example.com',
            'sophie.dubois@example.com',
            'paul.moreau@example.com'
        ]
    }
    
    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Save to Excel file
    df.to_excel('sample_contacts.xlsx', index=False)
    print("✅ Sample Excel file 'sample_contacts.xlsx' created successfully!")
    print("\nFile structure:")
    print("- genre: 'm', 'f', 'monsieur', 'madame' (gender for greeting)")
    print("- nom: Last name of the recipient")
    print("- email: Email address of the recipient")

if __name__ == "__main__":
    create_sample_excel()
