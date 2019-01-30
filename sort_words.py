infile = open("words_alpha.txt")
words = [x for x in infile]
words.sort()
infile.close()
outfile = open("sortedwords.txt","w")
for word in words:
    outfile.write(word)
