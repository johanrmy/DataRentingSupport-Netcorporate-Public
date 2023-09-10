from sqlalchemy import create_engine, Column, Integer, String, DateTime
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.declarative import declarative_base
import pandas as pd
from datetime import datetime
from colorama import Fore, init
init()
db = declarative_base()
class Ticket(db):
    __tablename__ = 'ticket'
    id = Column(Integer, primary_key=True, autoincrement=True)
    id_ticket = Column(Integer)
    serie = Column(String(40))
    proceso = Column(String(40))
    pedido = Column(String(40))
    descripcion = Column(String(50))
    asignacion = Column(String(40))
    cliente = Column(String(40))
    usuario = Column(String(40))
    prioridad = Column(String(40))
    etapa = Column(String(40))
    tipo = Column(String(40))
    email = Column(String(40))
    telefono = Column(String(40))
    descripcion_ticket = Column(String(400))
    cuenta = Column(String(40))
    creado = Column(DateTime)
    actualizado = Column(DateTime)
class DbManagement:
    def __init__(self, host, user, password, database):
        self.host = host
        self.user = user
        self.password = password
        self.database = database
        self.connection = None
        self.cursor = None
        self.engine = None
        try:
            self.engine = create_engine(f'mysql://{self.user}:{self.password}@{self.host}/{self.database}')
            db.metadata.create_all(self.engine)
            print(Fore.GREEN + "Base de datos creada exitosamente.")
        except Exception as e:
            print(Fore.RED + "Error al crear la base de datos:" + str(e))

    def connect(self):
        try:
            Session = sessionmaker(bind=self.engine)
            self.session = Session()
            self.session.query(Ticket).delete()
            print(Fore.GREEN + "\nConexión aceptada a " + Fore.CYAN + str(self.database))
        except Exception as e:
            print(Fore.RED + "\nError de conexión a la base de datos: " + str(e))

    def disconnect(self):
        if self.session:
            self.session.close()
            self.session = None
        if self.engine:
            self.engine.dispose()
            self.engine = None

    def save_data_from_excel(self, excel_file):
        try:
            excel_data = pd.read_excel(excel_file, header=0)
            data = excel_data.fillna(value="NULL")

            for index, row in data.iterrows():
                fecha_creado = datetime.strptime(row['Creado en'], '%d/%m/%Y %H:%M:%S')
                fecha_actualizado = datetime.strptime(row['Ultima Actualizacion'], '%d/%m/%Y %H:%M:%S')
                ticket = Ticket(
                    id_ticket=row['ID'],
                    serie=row['serie'],
                    proceso=row['proceso'],
                    pedido=row['pedido'],
                    descripcion=row['descripcion'],
                    asignacion=row['asignacion'],
                    cliente=row['cliente'],
                    usuario=row['usuario'],
                    prioridad=row['prioridad'],
                    etapa=row['etapa'],
                    tipo=row['tipo'],
                    email=row['email'],
                    telefono=row['telefono'],
                    descripcion_ticket=row['descripcion_ticket'],
                    cuenta=row['cuenta'],
                    creado=fecha_creado,
                    actualizado=fecha_actualizado
                )
                self.session.add(ticket)
            self.session.commit()
            print(Fore.GREEN + "Datos exportados exitosamente a la base de datos")
        except Exception as e:
            print(Fore.RED + "Error: Archivo xlsx no encontrado" + str(e))
