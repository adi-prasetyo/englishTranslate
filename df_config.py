from sqlalchemy import create_engine

def create_postgres_engine():
    return create_engine("postgresql+psycopg2://postgres:lalala@localhost:5432/postgres")


def create_aws_engine():
    return create_engine('postgresql+psycopg2://postgres:taji3030postgres@tajimaya-products-instance-1.crjzkl0txyjn.ap-northeast-1.rds.amazonaws.com:5432/postgres')