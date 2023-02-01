import random
import string


def create_word_search(words, size=40):
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

if __name__ == '__main__':
    words = ['dog', 'cat', 'bird', 'fish']
    grid = create_word_search(words)
    replace_dashes(grid)
    print_grid(grid)
    print_words_table(words)

