-- vybe-test: lua/debug_upvalues/debug_setupvalue_invalid_index_returns_nil
-- origin: languages/lua/tests/lua/test_debug_upvalues.rs

local __w1 = "nil"
local __i = 0

local function f() return 10 end
local res = debug.setupvalue(f, 2, 20)
do local __t = tostring(tostring(res)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
