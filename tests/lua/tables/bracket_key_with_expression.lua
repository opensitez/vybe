-- vybe-test: lua/tables/bracket_key_with_expression
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "7"
local __i = 0

local k = "x"
local t = {}
t[k] = 7
do local __t = tostring(t.x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
