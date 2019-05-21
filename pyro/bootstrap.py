# -*- coding: utf-8 -*-

def create_addin():
    try: 
        import pyro.addin as addin
        return addin.AddIn()
        
    except:
        import pyro.helpers
        pyro.helpers.exception_as_message()
