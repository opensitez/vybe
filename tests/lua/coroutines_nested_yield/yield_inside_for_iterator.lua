-- vybe-test: lua/coroutines_nested_yield/yield_inside_for_iterator
-- origin: languages/lua/tests/lua/test_coroutines_nested_yield.rs

local __w1 = "1234"
local __i = 0

local co = coroutine.create(function()
  for i = 1, 3 do coroutine.yield(i) end
  return 4
end)
local r = ""
while coroutine.status(co) ~= "dead" do
  local ok, val = coroutine.resume(co)
  if ok and val then r = r .. val end
end
do local __t = tostring(r); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
