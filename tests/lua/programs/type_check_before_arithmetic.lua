-- vybe-test: lua/programs/type_check_before_arithmetic
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "6"
local __i = 0

local x = "5"
if type(x) == "number" then do local __t = tostring(x + 1); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end else do local __t = tostring(tonumber(x) + 1); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
