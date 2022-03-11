import logging
class log_help:
    '''
    Logger Class with loging functions
    
    '''
    
    def __init__ (self):
        logging.basicConfig(filename = "log_file", level = logging.INFO, format = '%(asctime)s %(levelname)s %(message)s')
        
            
    def log(self, msg, _type= "info"):
        if (_type == "error"):
            logging.error(msg)
        elif (_type == "fatal"):
            logging.fatal(msg)
        else:
            logging.info(msg)
      
            
            
            