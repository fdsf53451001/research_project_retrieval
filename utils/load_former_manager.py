
def get_former_manager() -> list:
    with open('data/曾任委員.txt', 'r') as f:
        former_manager = f.read().split('\n')
    return former_manager