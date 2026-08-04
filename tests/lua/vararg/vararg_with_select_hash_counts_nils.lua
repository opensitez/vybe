-- vybe-test: lua/vararg/vararg_with_select_hash_counts_nils
-- origin: languages/lua/tests/lua/test_vararg.rs

local __w1 = "3"
local __i = 0

local function count(...) return select('#', ...) end
do local __t = tostring(count(1, nil, 3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
