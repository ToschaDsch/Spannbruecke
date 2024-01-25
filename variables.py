from dataclasses import dataclass


@dataclass
class Variables:
    begin: str = 'Begin'
    end: str = 'End'
    screen_bh = (1200, 400)


class Result:
    def __init__(self, x: float, y: float, n: int):
        self.x = x
        self.y = y
        self.n = n
