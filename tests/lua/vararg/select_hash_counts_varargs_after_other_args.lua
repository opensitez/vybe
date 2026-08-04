-- vybe-test: lua/vararg/select_hash_counts_varargs_after_other_args
-- origin: languages/lua/tests/lua/test_vararg.rs

local __w1 = "4"
local __i = 0

do local __t = tostring(select("#", 1, 2, 3, 4)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
