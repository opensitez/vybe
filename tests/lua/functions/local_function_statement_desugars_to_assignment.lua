-- vybe-test: lua/functions/local_function_statement_desugars_to_assignment
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "12"
local __i = 0

local function twice(x) return x * 2 end
do local __t = tostring(twice(6)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
