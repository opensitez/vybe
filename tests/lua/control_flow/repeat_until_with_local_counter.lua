-- vybe-test: lua/control_flow/repeat_until_with_local_counter
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "3"
local __i = 0

local tries = 0
repeat tries = tries + 1 until tries == 3
do local __t = tostring(tries); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
