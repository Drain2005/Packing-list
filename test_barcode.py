# test_barcode.py
import os

import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def test_barcode_generation():
    """Test simple de génération de code-barres"""
    test_codes = [
        "B25019-2A",
        "B25018-3C", 
        "B25018-2C",
        "B25019-2C",
        "TEST123",
        "123456"
    ]
    
    print("="*60)
    print("TEST DE GÉNÉRATION DE CODES-BARRES")
    print("="*60)
    
    for code in test_codes:
        print(f"\nTest du code: {code}")
        
        # Méthode 1: Code128
       
           
           
                
        
        
        # Méthode 2: Code39
        
            
        

if __name__ == "__main__":
    test_barcode_generation()
    print("\n" + "="*60)
    print("TEST TERMINÉ - Vérifiez les fichiers .png générés")
    print("="*60)
