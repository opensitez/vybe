-- vybe-test: lua/garbage_collection_advanced/test_gc_finalizer_order
-- origin: languages/lua/tests/lua/test_garbage_collection_advanced.rs

local __w1 = "true"
local __i = 0

local log = ''
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
do local __t = tostring(string.sub(log, 1, 1) == 'B'); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
