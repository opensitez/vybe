-- vybe-test: lua/base_print_multi/test_print_multiple_nested
-- origin: languages/lua/tests/lua/test_base_print_multi.rs

local __w1 = "aa\tbb"
local __i = 0

do local __t = tostring("aa") .. "\t" .. tostring("bb"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
