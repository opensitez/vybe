-- vybe-test: lua/closures/upvalue_survives_after_outer_function_returns
-- origin: languages/lua/tests/lua/test_closures.rs

local __w1 = "42"
local __i = 0

local function outer()
  local secret = 42
  return function() return secret end
end
local f = outer()
do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
