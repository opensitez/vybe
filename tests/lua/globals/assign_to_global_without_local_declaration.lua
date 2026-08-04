-- vybe-test: lua/globals/assign_to_global_without_local_declaration
-- origin: languages/lua/tests/lua/test_globals.rs

local __w1 = "2"
local __i = 0

count = 1
count = count + 1
do local __t = tostring(count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
