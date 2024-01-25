class Cable:
    def __init__(self, name: int, end: [] = None, begin: [] = None,
                 coordinate: [] = None, name_of_tab: str = None):
        self.name: int = name
        self.end: [] = None
        self.end_2: [] = None
        self.begin: [] = begin
        self.coordinate = [] if coordinate is None else coordinate
        self.i_0: int = 0
        self.i_n: int = 0
        self.selected = False
        self.color = (0, 0, 0)
        self.name_of_tab = name_of_tab

    def __str__(self):
        return (f'cable {self.name}  *******************************************************\n'
                f'begin {self.begin}, end {self.end} \n'
                f'coordinates \n'
                f'x {self.coordinate[0]} \n,'
                f'y {self.coordinate[1]} \n'
                f'{self.coordinate}')
