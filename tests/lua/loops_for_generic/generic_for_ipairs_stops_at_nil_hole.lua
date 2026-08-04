-- vybe-test: lua/loops_for_generic/generic_for_ipairs_stops_at_nil_hole
-- origin: languages/lua/tests/lua/test_loops_for_generic.rs

local __w1 = "30"
local __i = 0

local t = {10, 20, nil, 40}
local sum = 0
for _, v in ipairs(t) do sum = sum + v end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
