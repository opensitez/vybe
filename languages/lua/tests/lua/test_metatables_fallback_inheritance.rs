//! Prototype inheritance hierarchies using __index functions and tables (Lua 5.x §2.4)

lua_print! {
    inherit_gp => {
        "local gp = {name = \"gp\", age = 70}\nlocal parent = setmetatable({name = \"parent\"}, {__index = gp})\nlocal child = setmetatable({name = \"child\"}, {__index = parent})\nprint(child.name .. \",\" .. child.age)\n",
        "child,70"
    },
    inherit_func_fallback => {
        "local fallback = function(t, k) return \"fallback:\" .. k end\nlocal t = setmetatable({}, {__index = fallback})\nprint(t.foo)\n",
        "fallback:foo"
    },
    inherit_override => {
        "local parent = {val = 10}\nlocal child = setmetatable({val = 20}, {__index = parent})\nprint(child.val)\n",
        "20"
    },
    inherit_self_param => {
        "local proto = {\n  greet = function(self) return \"hello \" .. self.name end\n}\nlocal obj = setmetatable({name=\"obj\"}, {__index = proto})\nprint(obj:greet())\n",
        "hello obj"
    },
    inherit_rawset_override => {
        "local proto = {x = 1}\nlocal obj = setmetatable({}, {__index = proto})\nrawset(obj, \"x\", 2)\nprint(obj.x, proto.x)\n",
        "2 1"
    },
    inherit_dynamic_proto => {
        "local protoA = {x = 1}\nlocal protoB = {x = 2}\nlocal mt = {__index = protoA}\nlocal obj = setmetatable({}, mt)\nprint(obj.x)\nmt.__index = protoB\nprint(obj.x)\n",
        "1\n2"
    },
    inherit_cycle_fails => {
        "local t1 = {}\nlocal t2 = setmetatable({}, {__index = t1})\nsetmetatable(t1, {__index = t2})\nlocal ok, err = pcall(function() return t1.x end)\nprint(ok)\n",
        "false"
    },
}
