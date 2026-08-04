-- vybe-test: lua/coroutines_state_machine/coroutine_acc
-- origin: languages/lua/tests/lua/test_coroutines_state_machine.rs

local __w1 = "15"
local __i = 0

local co = coroutine.create(function()
  local acc = 0
  while true do
    local n = coroutine.yield(acc)
    if n == nil then return acc end
    acc = acc + n
  end
end)
coroutine.resume(co)  -- start
for _, v in ipairs({1,2,3,4,5}) do
  coroutine.resume(co, v)
end
local _, total = coroutine.resume(co, nil)
do local __t = tostring(total); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
