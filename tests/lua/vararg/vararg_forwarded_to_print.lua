-- vybe-test: lua/vararg/vararg_forwarded_to_print
-- origin: languages/lua/tests/lua/test_vararg.rs

local __w1 = "1\t2\t3"
local __i = 0

function show(...) do local __t = tostring(...); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end
show(1, 2, 3)

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
