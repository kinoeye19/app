import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BD_PATH = os.path.join(BASE_DIR, 'data.db')

def setup():
    print(f"프로젝트 루트: {BASE_DIR}")
    print(f"DB위치: {DB_PATH}")


if __name__ == '__main__':
    setup()



