
class Tasks():
    def __init__(self, query):
        self.query = query
    def wikipedia1(self):
        import wikipedia
        query = self.query.replace("wikipedia", '')
        results = wikipedia.summary(query, sentences=2)
        print(results)
        return results
   
