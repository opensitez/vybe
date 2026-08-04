-- vybe-test: lua/iteration/ipairs_stops_on_nil_slot
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "1"
local __i = 0

local t = {1, nil, 3}
local c = 0
for _ in ipairs(t) do c = c + 1 end
do local __t = tostring(c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
