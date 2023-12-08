
from scripts.core import run_project

def main(*args, **kwargs):
    """Данный проект получает заказы Кар-Тел в HTML формате, на его основе по шаблону создает АТП и АВР и формате .docx и в случае если есть смета, обьеденяет смету с ТАП и АВР"""
    
    
    run_project(*args, **kwargs)


if __name__ == "__main__":
    main()