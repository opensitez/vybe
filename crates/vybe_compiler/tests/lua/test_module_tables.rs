//! Table as namespace/module pattern — Lua's idiom for encapsulation (Lua 5.x)

lua_print! {
    module_greet => {
        "local M = {}\nfunction M.greet(name) return \"hello \" .. name end\nprint(M.greet(\"lua\"))\n",
        "hello lua"
    },
    module_private_upvalue => {
        "local function make_counter()\n  local n = 0\n  return {\n    inc = function() n = n + 1 end,\n    get = function() return n end,\n  }\nend\nlocal c = make_counter()\nc.inc(); c.inc(); c.inc()\nprint(c.get())\n",
        "3"
    },
    module_independent_state => {
        "local function new_stack()\n  local data = {}\n  return {\n    push = function(v) data[#data+1] = v end,\n    pop = function() local v = data[#data]; data[#data] = nil; return v end,\n    size = function() return #data end,\n  }\nend\nlocal a = new_stack()\nlocal b = new_stack()\na.push(1); a.push(2)\nb.push(9)\nprint(a.size() .. \",\" .. b.size())\n",
        "2,1"
    },
    module_constants => {
        "local Config = {max = 100, min = 0, name = \"app\"}\nprint(Config.max - Config.min .. \",\" .. Config.name)\n",
        "100,app"
    },
    module_nested_calls => {
        "local M = {}\nfunction M.double(n) return n * 2 end\nfunction M.quadruple(n) return M.double(M.double(n)) end\nprint(M.quadruple(3))\n",
        "12"
    },
    module_default_state => {
        "local M = {values = {}}\nfunction M.add(v) M.values[#M.values+1] = v end\nfunction M.sum()\n  local s = 0\n  for _, v in ipairs(M.values) do s = s + v end\n  return s\nend\nM.add(5); M.add(10); M.add(15)\nprint(M.sum())\n",
        "30"
    },
    module_alias => {
        "local MyMath = {}\nfunction MyMath.sq(n) return n * n end\nlocal sq = MyMath.sq\nprint(sq(7))\n",
        "49"
    },
    module_nested_tables => {
        "local App = {db = {}, ui = {}}\nfunction App.db.query(q) return \"result:\" .. q end\nprint(App.db.query(\"users\"))\n",
        "result:users"
    },
}
