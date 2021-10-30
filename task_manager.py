import tdprocessor as tp
import os

def wikipedia1(query):
    import wikipedia
    query = query.replace("wikipedia", '')
    results = wikipedia.summary(query, sentences=2)
    print(results)
    return results
   
def openapp(query):
    ''' Opens apps based on query'''
    possible_apps = []
    TASK_DATA = tp.taskreader("applink")
    ambigious_path = []
    for dict in TASK_DATA:
        if query in list(dict.keys()) [0]:
            possible_apps.append(dict)
    if len(possible_apps) == 1:
        for i in possible_apps:
            print(list(i.keys()) [0],"@@@@@", list(i.values()) [0] [0])
        os.startfile(list(i.values()) [0] [0])
    elif len(possible_apps) > 1:
        print(f"Multiple programs found with {query} in it!")
        for i in possible_apps:
            print(list(i.keys()) [0],"@@@@@", list(i.values()) [0] [0])
    else:
        print("No such task found. Do you wanna add it if its present?")
        #tp.taskwriter()

        
                #os.startfile(path)
    
    #print(no_possible_apps, possible_apps)
        


#openapp('mp3tagsfortracks')

