import random
from termcolor import colored #pip install termcolor

guessed_words = set()

def generate_word():
    # Generate a secret word for the player to guess
    with open("five_letter_words.txt", "r") as file:
        words = file.readlines()
        return random.choice(words).strip().lower()

# Check if guess is correct
def check_guess(secret_word, guess):
    global game_over
    global invalid 
    global guessed_words 
    
    # Check length
    if len(guess) != len(secret_word):
        invalid = True 
        print("Invalid guess. Please try a 5 letter word.")
        return

    # Check if they are all letters
    if not guess.isalpha():
        invalid = True
        print("Invalid guess. Please enter only letters.")
        return
    
    # Check if it is a valid word
    with open("five_letter_words.txt", 'r') as file:
        word_list = file.read().splitlines()
        if guess not in word_list:
            invalid = True
            print("Invalid guess. '{}' is not a valid word.".format(guess))

    # Check if the word has already been guessed
    if guess in guessed_words:
        invalid = True
        print("This word has already been guessed. Try again.")
    else:
        guessed_words.add(guess)

    feedback = ""
    word_dict = {letter: secret_word.count(letter) for letter in secret_word}
    for i in range(5):
        if guess[i] == secret_word[i]:
            feedback += colored(guess[i], 'green')  # Correct letter in correct position
            word_dict[guess[i]] -= 1
        elif guess[i] in secret_word and word_dict[guess[i]] > 0:
            feedback += colored(guess[i], 'yellow')  # Correct letter in wrong position
            word_dict[guess[i]] -= 1
        else:
            feedback += guess[i]  # Incorrect letter
    if guess == secret_word:
        print("Congratulations! You guessed the word '{}' correctly.".format(secret_word))
        game_over = True

    return feedback

def play_wordle():
    # Main function
    global game_over
    global invalid
    global guessed_words
    
    while True: 
        print("\nWelcome to Wordle!")
        secret_word = generate_word()
        attempts = 6
        game_over = False
        invalid = False 

        while attempts > 0 and not game_over:
            guess = input("\nEnter your guess ({} attempts remaining): ".format(attempts)).lower()
            result = check_guess(secret_word, guess)
            if invalid is False:
                attempts -= 1
                print(result)
            invalid = False

        if attempts == 0 and not game_over:
            print("Sorry, you've run out of attempts. The secret word was '{}'.".format(secret_word))
            game_over = True
        
        if game_over: 
            play_again = input("\nDo you want to play again? (Y/N): ").strip().lower()
            if play_again != 'y':
                break
            else:
                guessed_words = set()

# Start the game
play_wordle()
