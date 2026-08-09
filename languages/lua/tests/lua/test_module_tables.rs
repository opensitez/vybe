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
module_with_init_function_sets_state => {
    "local M = {ready = false}\nfunction M.init() M.ready = true; M.value = 42 end\nM.init()\nprint(tostring(M.ready) .. ',' .. M.value)\n",
    "true,42"
},
module_lazy_singleton_initialized_on_first_call => {
    "local instance = nil\nlocal M = {}\nfunction M.get()\n  if not instance then instance = {count = 0} end\n  instance.count = instance.count + 1\n  return instance\nend\nM.get(); M.get()\nprint(M.get().count)\n",
    "3"
},
module_extends_another_by_copying_methods => {
    "local Base = {greet = function() return 'hi' end}\nlocal Child = {}\nfor k, v in pairs(Base) do Child[k] = v end\nfunction Child.farewell() return 'bye' end\nprint(Child.greet() .. ',' .. Child.farewell())\n",
    "hi,bye"
},
module_method_chain_returns_self => {
    "local Builder = {parts = {}}\nfunction Builder:add(s) self.parts[#self.parts+1] = s; return self end\nfunction Builder:build() return table.concat(self.parts, '-') end\nprint(Builder:add('a'):add('b'):add('c'):build())\n",
    "a-b-c"
},
module_version_field_accessible => {
    "local M = {_VERSION = '1.0.0', name = 'mymod'}\nprint(M.name .. '@' .. M._VERSION)\n",
    "mymod@1.0.0"
},
module_public_api_via_explicit_table => {
    "local function create_module()\n  local private = 'secret'\n  return {\n    public = function() return 'pub:' .. private end\n  }\nend\nlocal mod = create_module()\nprint(mod.public())\n",
    "pub:secret"
},
module_local_alias_for_nested_function => {
    "local M = {}\nlocal math_floor = math.floor\nfunction M.truncate(x) return math_floor(x) end\nprint(M.truncate(3.9))\n",
    "3"
} }
