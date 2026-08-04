-- vybe-test: lua/loops_for_generic_ipairs_mut/test_loops_for_generic_ipairs_mut_build_suffix
-- origin: languages/lua/tests/lua/test_loops_for_generic_ipairs_mut.rs

local __w1 = "1:1;2:2;"
local __i = 0

local out = ''
for i, value in ipairs({1,2}) do out = out .. tostring(i) .. ':' .. tostring(value) .. ';' end
do local __t = tostring(out); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
