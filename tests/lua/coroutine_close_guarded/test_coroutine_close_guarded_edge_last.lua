-- vybe-test: lua/coroutine_close_guarded/test_coroutine_close_guarded_edge_last
-- origin: languages/lua/tests/lua/test_coroutine_close_guarded.rs

local __w1 = "true"
local __i = 0

local co = coroutine.create(function()
  local x = 18
  return x
end)
coroutine.resume(co)
do local __t = tostring(pcall(coroutine.close, co) ~= nil); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
