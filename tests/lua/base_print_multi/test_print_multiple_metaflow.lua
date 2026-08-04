-- vybe-test: lua/base_print_multi/test_print_multiple_metaflow
-- origin: languages/lua/tests/lua/test_base_print_multi.rs

local __w1 = "u\t7\t8\t9"
local __i = 0

do local __t = tostring("u") .. "\t" .. tostring(7) .. "\t" .. tostring(8) .. "\t" .. tostring(9); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
