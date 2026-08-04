-- vybe-test: lua/loops_for_generic/generic_for_coroutine_producer_consumer
-- origin: languages/lua/tests/lua/test_loops_for_generic.rs

local __w1 = "10,20,30,"
local __i = 0

local function producer(t)
  local i = 0
  return function()
    i = i + 1
    return t[i]
  end
end
local s = ''
for v in producer({10, 20, 30}) do s = s .. v .. ',' end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
