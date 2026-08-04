-- vybe-test: lua/coroutine_patterns/coroutine_communication
-- origin: languages/lua/tests/lua/test_coroutine_patterns.rs

local __w1 = "10\t15\t18"
local __i = 0

local co = coroutine.create(function(start)
  local acc = start
  while true do
    local n = coroutine.yield(acc)
    if n == nil then break end
    acc = acc + n
  end
  return acc
end)
local _, v1 = coroutine.resume(co, 10)
local _, v2 = coroutine.resume(co, 5)
local _, v3 = coroutine.resume(co, 3)
do local __t = tostring(v1) .. "\t" .. tostring(v2) .. "\t" .. tostring(v3); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
