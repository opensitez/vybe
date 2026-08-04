-- vybe-test: lua/vararg/varargs_packed_into_table_spread
-- origin: languages/lua/tests/lua/test_vararg.rs

local __w1 = "9"
local __i = 0

function all(...) return {...} end
local t = all(4, 5)
do local __t = tostring(t[1] + t[2]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
