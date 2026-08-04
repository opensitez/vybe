-- vybe-test: lua/programs/simulate_state_machine_two_states
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "done"
local __i = 0

local state = "idle"
if state == "idle" then state = "run" end
if state == "run" then state = "done" end
do local __t = tostring(state); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
