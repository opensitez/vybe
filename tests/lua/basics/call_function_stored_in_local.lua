-- vybe-test: lua/basics/call_function_stored_in_local
-- origin: languages/lua/tests/lua/test_basics.rs

local __w1 = "3"
local __i = 0

local function add(a, b) return a + b end
do local __t = tostring(add(1, 2)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
