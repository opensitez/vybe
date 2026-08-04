-- vybe-test: lua/string_find_first/test_string_find_first_metaflow
-- origin: languages/lua/tests/lua/test_string_find_first.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(string.find("alpha0 bravo1 charlie2 delta3 echo4 foxtrot5 golf6 hotel7 india8 juliet9 kilo10 lima11 mike12 november13 oscar14 papa15 quebec16 romeo17 sierra18 tango19 uniform20", "lima11", 1, true) == 81); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
