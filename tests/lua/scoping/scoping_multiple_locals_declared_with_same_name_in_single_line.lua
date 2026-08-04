-- vybe-test: lua/scoping/scoping_multiple_locals_declared_with_same_name_in_single_line
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "200"
local __i = 0

local x, x = 100, 200
do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
