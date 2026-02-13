import os
import datetime
from contextlib import contextmanager
from sqlalchemy import Column, String, Text, DateTime, create_engine
from sqlalchemy.orm import declarative_base, sessionmaker
from dotenv import load_dotenv

# === Load .env dari root project ===
dotenv_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), ".env")
load_dotenv(dotenv_path)

Base = declarative_base()

class ZoneData(Base):
    __tablename__ = 'zone_data'
    zone = Column(String, primary_key=True)
    data = Column(Text, nullable=False)
    updated_at = Column(DateTime, default=datetime.datetime.utcnow, onupdate=datetime.datetime.utcnow)

def get_engine():
    db_url = os.getenv("DATABASE_URL")
    if not db_url:
        raise RuntimeError("DATABASE_URL belum didefinisikan di .env")
    return create_engine(db_url, echo=False, future=True)

@contextmanager
def get_session():
    engine = get_engine()
    Session = sessionmaker(bind=engine, autoflush=False, autocommit=False)
    session = Session()
    try:
        yield session
        session.commit()
    except:
        session.rollback()
        raise
    finally:
        session.close()

def create_tables():
    engine = get_engine()
    Base.metadata.create_all(engine)