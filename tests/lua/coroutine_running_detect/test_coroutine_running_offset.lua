-- vybe-test: lua/coroutine_running_detect/test_coroutine_running_offset
-- origin: languages/lua/tests/lua/test_coroutine_running_detect.rs

local __w1 = "true"
local __i = 0

local inside = false
local co = coroutine.create(function()
  inside = coroutine.running() ~= nil
end)
do local __t = tostring(coroutine.resume(co)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
if inside then do local __t = tostring(true); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end else do local __t = tostring(false); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
