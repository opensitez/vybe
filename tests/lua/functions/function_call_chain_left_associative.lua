-- vybe-test: lua/functions/function_call_chain_left_associative
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "9"
local __i = 0

local function id(x) return x end
do local __t = tostring(id(id(id(9)))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
