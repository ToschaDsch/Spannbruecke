from dataclasses import dataclass


@dataclass
class Variables:
    begin: str = 'Begin'
    end: str = 'End'
    screen_bh = (1200, 400)
