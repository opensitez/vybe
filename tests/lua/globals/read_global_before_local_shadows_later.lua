-- vybe-test: lua/globals/read_global_before_local_shadows_later
-- origin: languages/lua/tests/lua/test_globals.rs

local __w1 = "1"
local __i = 0

baz = 5
local baz = 1
do local __t = tostring(baz); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
