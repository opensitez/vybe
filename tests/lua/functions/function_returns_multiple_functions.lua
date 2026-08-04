-- vybe-test: lua/functions/function_returns_multiple_functions
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "3"
local __i = 0

local function make_pair()
  return function() return 1 end, function() return 2 end
end
local a, b = make_pair()
do local __t = tostring(a() + b()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
