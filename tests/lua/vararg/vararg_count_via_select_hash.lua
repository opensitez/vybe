-- vybe-test: lua/vararg/vararg_count_via_select_hash
-- origin: languages/lua/tests/lua/test_vararg.rs

local __w1 = "3"
local __i = 0

function n(...) return select("#", ...) end
do local __t = tostring(n("a", "b", "c")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
