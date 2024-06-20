a = "hello"
print(a)
b=a
c="hello"
print(b)
print(a is b)
print(a is c)  #这种情况下，a和c指向的是同一个对象，所以返回True，如果在java中，这种情况下返回false，若是String等特殊类型会在源码中重写equals方法，所以会返回true
name = "ada"
print(name.title())
print(name.upper())
print(name.lower())
print(2/3)
print(2.0/3)
print("2"+name)