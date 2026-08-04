-- vybe-test: lua/base_print_multi/test_print_multiple_prefixed
-- origin: languages/lua/tests/lua/test_base_print_multi.rs

local __w1 = "m\tn\to"
local __i = 0

do local __t = tostring("m") .. "\t" .. tostring("n") .. "\t" .. tostring("o"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
