-- vybe-test: lua/base_print_multi/test_print_multiple_simple
-- origin: languages/lua/tests/lua/test_base_print_multi.rs

local __w1 = "b\t2\tfalse"
local __i = 0

do local __t = tostring("b") .. "\t" .. tostring(2) .. "\t" .. tostring(false); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
