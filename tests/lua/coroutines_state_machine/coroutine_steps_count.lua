-- vybe-test: lua/coroutines_state_machine/coroutine_steps_count
-- origin: languages/lua/tests/lua/test_coroutines_state_machine.rs

local __w1 = "5"
local __i = 0

local steps = 0
local co = coroutine.create(function()
  for i = 1, 5 do
    steps = steps + 1
    coroutine.yield()
  end
end)
while coroutine.status(co) ~= "dead" do
  coroutine.resume(co)
end
do local __t = tostring(steps); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
