-- vybe-test: lua/generic_for_protocol/iterator_coroutine_wrapper
-- origin: languages/lua/tests/lua/test_generic_for_protocol.rs

local __w1 = "10"
local __i = 0

local function gen(n)
  return coroutine.wrap(function()
    for i = 1, n do coroutine.yield(i) end
  end)
end
local s = 0
for v in gen(4) do s = s + v end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
