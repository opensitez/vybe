-- vybe-test: lua/vararg/select_returns_nth_vararg
-- origin: languages/lua/tests/lua/test_vararg.rs

local __w1 = "30\t40"
local __i = 0

function third(...) return select(3, ...) end
do local __t = tostring(third(10, 20, 30, 40)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
