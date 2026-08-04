-- vybe-test: lua/literals/nil_in_table_field_read
-- origin: languages/lua/tests/lua/test_literals.rs

local __w1 = "nil"
local __i = 0

local t = {x = nil}
do local __t = tostring(tostring(t.x)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
