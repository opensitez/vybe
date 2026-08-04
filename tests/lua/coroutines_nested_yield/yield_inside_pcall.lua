-- vybe-test: lua/coroutines_nested_yield/yield_inside_pcall
-- origin: languages/lua/tests/lua/test_coroutines_nested_yield.rs

local __w1 = "yielded"
local __i = 0

local co = coroutine.create(function()
  local ok, err = pcall(function()
    coroutine.yield("yielded")
  end)
  return ok, err
end)
local _, val = coroutine.resume(co)
do local __t = tostring(val); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
