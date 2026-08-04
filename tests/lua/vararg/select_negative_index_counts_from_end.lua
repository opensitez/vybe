-- vybe-test: lua/vararg/select_negative_index_counts_from_end
-- origin: languages/lua/tests/lua/test_vararg.rs

local __w1 = "9"
local __i = 0

function last(...) return select(-1, ...) end
do local __t = tostring(last(1, 2, 9)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
