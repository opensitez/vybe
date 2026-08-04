-- vybe-test: lua/scoping/upvalue_from_function_param_survives_return
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "55"
local __i = 0

local function make(n)
  return function() return n end
end
local f = make(55)
do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
