-- vybe-test: lua/debug_upvalues/debug_setupvalue_returns_name_of_upvalue
-- origin: languages/lua/tests/lua/test_debug_upvalues.rs

local __w1 = "my_upval"
local __i = 0

local my_upval = 10
local function f() return my_upval end
local name = debug.setupvalue(f, 1, 20)
do local __t = tostring(name); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
