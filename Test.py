
import operator
from __builtin__ import sorted

data = dict()

data["abc"] = 1
data["tef"] = 9
data["smth"] = 9

data["org"] = 6

print data

dataList = sorted(data.iteritems(),key=operator.itemgetter(1,0),reverse=True)

print dataList


table = {}

table["c1"] = [[1,2,3]]
table["c2"] = [[4,5,6]]
table["c5 c4 (%)"] = []

table.get("c5 c4 (%)").append([7,8,9])

c5c4data = table["c5 c4 (%)"]
print "c5c4data = ", c5c4data

print table


a = [1,2,3,4,5,23]

average = float(sum(a)) / len(a)
print average

table2 = {}
table2["c1"] = [{"key1":1}, {"key1":2}]
print table2


def getKey(item):
    return item[0]
data = [["c5", 1], ["c3", 2], ["c4", 8]]

data.sort(cmp=None, key=getKey, reverse=True)
print data

a = 3
b = "Hello B"
str1 = "This is a message: [%s:%d]" % (b, a)


print str1




def addValueToDict(key, table, item):
    origValues = table[key]
    origValues.append(item)
    
    table[key] = origValues
    
    return table
if __name__ == '__main__':
    key = "key1"
    myTable = {}
    myTable[key] = ["item1", "item2"]
    
    print myTable
    addValueToDict(key, myTable, "item3")
    print myTable
    pass