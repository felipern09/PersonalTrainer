from sqlalchemy import create_engine
from sqlalchemy import MetaData
from sqlalchemy.orm import declarative_base
from sqlalchemy import Column, Integer, Float, String, ForeignKey, Boolean
from sqlalchemy.orm import relationship
from sk import sql

engine = create_engine(sql, echo=True, future=True)
metadata_obj = MetaData()
Base = declarative_base()


class Personal(Base):
    __tablename__ = "personal"
    id = Column(Integer, primary_key=True)
    nome = Column(String, nullable=False)
    email = Column(String, nullable=False)
    whatsapp = Column(String, default="61")
    tipo_personal = Column(String, nullable=False)
    status = Column(String, nullable=False)
    aulas = relationship('Aulas', backref='professor')


class Aulas(Base):
    __tablename__ = "aulas"
    idaulas = Column(Integer, primary_key=True)
    personal = Column(Integer, ForeignKey('personal.id'))
    mes = Column(String, nullable=False)
    simples1 = Column(Integer, default=0)
    simples2 = Column(Integer, default=0)
    dupla1 = Column(Integer, default=0)
    dupla2 = Column(Integer, default=0)
    tripla1 = Column(Integer, default=0)
    tripla2 = Column(Integer, default=0)
    total_aulas = Column(Integer, default=0)
    descaula = Column(Integer, default=0)
    descvalor = Column(Float, default=0)
    acrescaula = Column(Integer, default=0)
    acrescvalor = Column(Float, default=0)
    valortotalemdia = Column(Integer, default=0)
    valortotalatraso = Column(Integer, default=0)
    valorcobrado = Column(Float, default=0)
    foipago = Column(Boolean, default=False)
    valorpago = Column(Float, default=0)
    credito = Column(Float, default=0)
    debito = Column(Float, default=0)


class Valores(Base):
    __tablename__ = "valores"
    id = Column(Integer, primary_key=True)
    internosimplesnodesc = Column(Float, nullable=False)
    internoduplanodesc = Column(Float, nullable=False)
    internotriplanodesc = Column(Float, nullable=False)
    externosimplesnodesc = Column(Float, nullable=False)
    externoduplanodesc = Column(Float, nullable=False)
    externotriplanodesc = Column(Float, nullable=False)
    internosimples1a10 = Column(Float, nullable=False)
    internosimples11a30 = Column(Float, nullable=False)
    internosimples31a50 = Column(Float, nullable=False)
    internosimples51a100 = Column(Float, nullable=False)
    internosimples101a120 = Column(Float, nullable=False)
    internosimplesacima120 = Column(Float, nullable=False)
    internodupla1a60 = Column(Float, nullable=False)
    internodupla61a119 = Column(Float, nullable=False)
    internoduplaacima119 = Column(Float, nullable=False)
    internotripla1a60 = Column(Float, nullable=False)
    internotripla61a119 = Column(Float, nullable=False)
    internotriplaacima119 = Column(Float, nullable=False)
    externosimples1a10 = Column(Float, nullable=False)
    externosimples11a30 = Column(Float, nullable=False)
    externosimples31a50 = Column(Float, nullable=False)
    externosimples51a100 = Column(Float, nullable=False)
    externosimples101a120 = Column(Float, nullable=False)
    externosimplesacima120 = Column(Float, nullable=False)
    externodupla1a60 = Column(Float, nullable=False)
    externodupla61a119 = Column(Float, nullable=False)
    externoduplaacima119 = Column(Float, nullable=False)
    externotripla1a60 = Column(Float, nullable=False)
    externotripla61a119 = Column(Float, nullable=False)
    externotriplaacima119 = Column(Float, nullable=False)


class Usuario(Base):
    __tablename__ = "usuario"
    id = Column(Integer, primary_key=True)
    nome = Column(String, nullable=False)
    email = Column(String, nullable=False)
    senha = Column(String, nullable=False)
    servidor = Column(String, nullable=False)
    porta = Column(Integer, default=0)
    assinatura = Column(String, nullable=False)


Base.metadata.create_all(engine)
