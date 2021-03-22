import re
# taking input fron user for the no. of words
saw = int(input("enter the no. to search for words: \n"))
# class for searching the word


class search_w:
    def __init__(self):               # method for mandatory inputs
        # variable for opening the file
        self.file_new = open("input.txt",'r')
        # making the file to string using .read method
        self.filter_str = self.file_new.read()
        # taking word from user to search
        self.word_search = input("enter the word:\n")
        # used for saving the output file in.txt file
        self.file_save = self.word_search + '.txt'
        # opening the output file to save with new lines
        self.write_file = open(self.file_save, 'w+')

    def spilt(self):          # method for spilting the file
        # spilting the prev filter_str for making each line into list
        file_sp = self.filter_str.split()
        # used for counting the repearing of word
        count = 0
        # done iteration for printing the lines into new file
        for i in range(len(file_sp)):
            # matching the word per each line
            word_search1 = re.match(self.word_search, file_sp[i], re.I | re.M)
            # for checking word is present or not
            if word_search1:

                count += 1
   
                # for printing the prev word and next word of the desired word
                str78 = (file_sp[i-1] + ' ' + file_sp[i] + ' ' + file_sp[i+1])
                self.write_file.write(str(count)+':')
                self.write_file.write(str(str78) + '\n')
                # for getting the no. of word repeated
        self.write_file.write('no. of times word comes:' + str(count) + '\n')
        print(count)
        for i in range(saw):     # to take multiple output from the user
        # variable declared for the class
        sample = search_w()
    # calling the spilt function using the variable
    sample.spilt()