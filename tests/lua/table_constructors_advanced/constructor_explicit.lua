-- vybe-test: lua/table_constructors_advanced/constructor_explicit
-- origin: languages/lua/tests/lua/test_table_constructors_advanced.rs

local __w1 = "ten"
local __i = 0

local t = {[10]="ten", [20]="twenty"}
do local __t = tostring(t[10]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
