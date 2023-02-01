import random
import string
import cups
import re

def create_word_search(words, size=30):
    """Creates a word search from a list of words and returns the grid as a list of lists."""
    grid = [['-' for _ in range(size)] for _ in range(size)]
    
    # Function to check if word fits in the grid
    def word_fits(word, x, y, direction):
        for i, letter in enumerate(word):
            if x < 0 or x >= size or y < 0 or y >= size:
                return False
            if grid[x][y] != '-' and grid[x][y] != letter:
                return False
            if direction == 'down':
                x += 1
            elif direction == 'right':
                y += 1
            elif direction == 'diagonal':
                x += 1
                y += 1
            elif direction == 'backward_horizontal':
                y -= 1
            elif direction == 'backward_vertical':
                x -= 1
            elif direction == 'backward_diagonal':
                x -= 1
                y -= 1
        return True
    
    for word in words:
        while True:
            direction = random.choice(['down', 'right', 'diagonal', 'backward_horizontal', 'backward_vertical', 'backward_diagonal'])
            x = random.randint(0, size-1)
            y = random.randint(0, size-1)
            if word_fits(word, x, y, direction):
                break
        for i, letter in enumerate(word):
            grid[x][y] = letter
            if direction == 'down':
                x += 1
            elif direction == 'right':
                y += 1
            elif direction == 'diagonal':
                x += 1
                y += 1
            elif direction == 'backward_horizontal':
                y -= 1
            elif direction == 'backward_vertical':
                x -= 1
            elif direction == 'backward_diagonal':
                x -= 1
                y -= 1
    
    return grid

def replace_dashes(grid):
    x=0
    for row in grid:
        y=0
        for character in row:
            grid[x][y]=random.choice(string.ascii_letters).lower()
            y+=1
        x+=1

def print_grid(grid):
    """Prints the grid as a string."""
    for row in grid:
        print(' '.join(row))

def print_words_table(words):
    """Prints the words in a table format."""
    print('\nWords:')
    for i, word in enumerate(words):
        print(f'{i+1}. {word}')

def select_printer():
    """Lists available CUPS printers and allows the user to select one."""
    conn = cups.Connection()
    printers = conn.getPrinters()
    print("Available printers:")
    for i, printer in enumerate(printers):
        print("{}: {}".format(i, printer))
    selected_printer = int(input("Select a printer by its number: "))
    return list(printers.keys())[selected_printer]

def print_to_selected_printer(grid, printer_id):
    """Prints the word search puzzle to the selected printer."""
    conn = cups.Connection()
    grid_string = '\n'.join([' '.join(row) for row in grid])
    with open("/tmp/print.txt", "w") as f:
        f.write(grid_string)
    conn.printFile(printer_id, "/tmp/print.txt", "Print Job", {})


if __name__ == '__main__':
    with open("words.txt", "r") as f:
        words=f.read().split("\n")
    original_words=words.copy()
    original_words.remove(original_words[0])
    original_words.remove("")
    counter=0
    for x in words:
        if counter==0:
            title=words[counter]
            words.remove(words[0])
        else:
            words[counter] = re.sub(r'[^a-z]+', '', x, flags=re.IGNORECASE).lower()
        counter+=1
    words.remove("")
    grid = create_word_search(words)
    replace_dashes(grid)
    print_grid(grid)
    print_words_table(words)
    selected_printer = select_printer()
    print_to_selected_printer(grid, selected_printer)
