-- vybe-test: lua/iteration/ipairs_iterator_call_returns_index_and_value
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "1,10"
local __i = 0

local t = {10, 20}
local i, v = ipairs(t)()
do local __t = tostring(i .. "," .. v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
