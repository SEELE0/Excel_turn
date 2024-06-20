pizzas=['A','B','C']
for pizza in pizzas:
    print(pizza+'我不喜欢! \n')
print('I like pepperoni pizza.')

integer = [test **3  for test in range(1,11)]
print(integer)
#列表复制
integer2 = integer[:]
#直接赋值 指向同一个列表
integer3 = integer
integer.append(15)
integer2.append(18)
integer3.append(20)
print(integer)
print(integer2)
print(integer3) 
