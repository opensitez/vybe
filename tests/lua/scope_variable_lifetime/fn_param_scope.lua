-- vybe-test: lua/scope_variable_lifetime/fn_param_scope
-- origin: languages/lua/tests/lua/test_scope_variable_lifetime.rs

local __w1 = "nil"
local __i = 0

local function f(a) return a + 1 end
f(5)
do local __t = tostring(tostring(a)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
