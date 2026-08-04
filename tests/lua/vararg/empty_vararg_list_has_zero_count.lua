-- vybe-test: lua/vararg/empty_vararg_list_has_zero_count
-- origin: languages/lua/tests/lua/test_vararg.rs

local __w1 = "0"
local __i = 0

function n(...) return select("#", ...) end
do local __t = tostring(n()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
