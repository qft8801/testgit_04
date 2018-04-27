def PrintAllParagraphs(doc):
    count = doc.Paragraphs.Count
    for i in range(count - 1, -1, -1):
        pr = doc.Paragraphs[i].Range
        print()
        pr.Text