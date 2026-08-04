-- vybe-test: lua/coroutines_state_machine/coroutine_state_transition
-- origin: languages/lua/tests/lua/test_coroutines_state_machine.rs

local __w1 = "idle,running,stopped"
local __i = 0

local function sm()
  coroutine.yield("idle")
  coroutine.yield("running")
  coroutine.yield("stopped")
end
local co = coroutine.create(sm)
local states = {}
for _ = 1, 3 do
  local _, s = coroutine.resume(co)
  states[#states+1] = s
end
do local __t = tostring(table.concat(states, ",")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
