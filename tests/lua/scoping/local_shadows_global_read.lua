-- vybe-test: lua/scoping/local_shadows_global_read
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "2"
local __i = 0

x = 1
local x = 2
do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
