//! Object-oriented class and inheritance patterns via metatables (Lua 5.x)

lua_print! {
    oop_class_instantiation => {
        "local Point = {}\nPoint.__index = Point\nfunction Point.new(x, y)\n  return setmetatable({x=x, y=y}, Point)\nend\nlocal p = Point.new(3, 4)\nprint(p.x, p.y)\n",
        "3\t4"
    },
    oop_class_methods => {
        "local Point = {}\nPoint.__index = Point\nfunction Point.new(x, y)\n  return setmetatable({x=x, y=y}, Point)\nend\nfunction Point:sum() return self.x + self.y end\nlocal p = Point.new(10, 20)\nprint(p:sum())\n",
        "30"
    },
    oop_inheritance => {
        "local Animal = {}\nAnimal.__index = Animal\nfunction Animal:speak() return \"sound\" end\nlocal Dog = setmetatable({}, Animal)\nDog.__index = Dog\nfunction Dog:speak() return \"woof\" end\nlocal d = setmetatable({}, Dog)\nprint(d:speak())\n",
        "woof"
    },
    oop_inheritance_fallback => {
        "local Animal = {}\nAnimal.__index = Animal\nfunction Animal:speak() return \"sound\" end\nlocal Dog = setmetatable({}, Animal)\nDog.__index = Dog\nlocal d = setmetatable({}, Dog)\nprint(d:speak())\n",
        "sound"
    },
    oop_self_methods => {
        "local Account = {balance = 0}\nAccount.__index = Account\nfunction Account:deposit(v) self.balance = self.balance + v end\nlocal a = setmetatable({}, Account)\na:deposit(100)\nprint(a.balance)\n",
        "100"
    },
    oop_dynamic_dispatch => {
        "local ClassA = {name = \"A\"}\nClassA.__index = ClassA\nlocal ClassB = {name = \"B\"}\nClassB.__index = ClassB\nlocal function get_obj(cond)\n  if cond then return setmetatable({}, ClassA) else return setmetatable({}, ClassB) end\nend\nprint(get_obj(true).name, get_obj(false).name)\n",
        "A\tB"
    },
}
