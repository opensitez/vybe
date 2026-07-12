lua_print! {
    test_gc_weak_tables_keys => {
        "local t = {}
setmetatable(t, {__mode = 'k'})
local k = {}
t[k] = 1
k = nil
collectgarbage('collect')
local count = 0
for k, v in pairs(t) do count = count + 1 end
print(count)",
        "0"
    },
    test_gc_weak_tables_values => {
        "local t = {}
setmetatable(t, {__mode = 'v'})
local v = {}
t[1] = v
v = nil
collectgarbage('collect')
local count = 0
for k, val in pairs(t) do count = count + 1 end
print(count)",
        "0"
    },
    test_gc_finalizer_resurrection => {
        "local resurrected
local t = {}
setmetatable(t, {__gc = function(o) resurrected = o end})
t = nil
collectgarbage('collect')
print(type(resurrected))",
        "table"
    },
    test_gc_cycle_collection => {
        "local a = {}
local b = {}
a.b = b
b.a = a
local mt = {__mode = 'k'}
local weak = setmetatable({}, mt)
weak[a] = 1
weak[b] = 2
a = nil
b = nil
collectgarbage('collect')
local count = 0
for k, v in pairs(weak) do count = count + 1 end
print(count)",
        "0"
    },
    test_gc_step_execution => {
        "local pre = collectgarbage('count')
local t = {}
for i = 1, 10000 do t[i] = tostring(i) end
t = nil
local b = collectgarbage('step')
local post = collectgarbage('count')
print(type(b) == 'boolean')",
        "true"
    },
    test_gc_finalizer_order => {
        "local log = ''
local function make_obj(name, parent)
    local o = {name = name}
    setmetatable(o, {__gc = function(self) log = log .. self.name end})
    if parent then parent.child = o end
    return o
end
do
    local a = make_obj('A')
    local b = make_obj('B', a)
end
collectgarbage('collect')
print(string.sub(log, 1, 1) == 'B')",
        "true"
    }
}
