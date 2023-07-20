from pathlib import Path

class Configuration:
    def __init__(self, test: bool=False):
        _server = r'mssqldb2t.ad.vse.sk\BLUEPRISM' if test else r'mssqldb2.ad.vse.sk\BLUEPRISM'
        self.connection_string = f"""DRIVER={{SQL Server}};
            Server={_server};
            Database=robot_run_log;
            Trusted_Connection=yes;"""
        self.working_folder = Path(r'\\ad.vse.sk\APP\aplikacie\Blue_Prism\4. Reporting\robot_run_log')
        self.error_filename = 'robot_run_log_errors_t.txt' if test else 'robot_run_log_errors.txt' 
        self.error_path = self.working_folder.joinpath(self.error_filename)