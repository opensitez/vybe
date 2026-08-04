-- vybe-test: lua/loops_for_generic/generic_for_using_next_as_iterator
-- origin: languages/lua/tests/lua/test_loops_for_generic.rs

local __w1 = "6"
local __i = 0

local t = {a=1, b=2, c=3}
local count = 0
for k, v in pairs(t) do count = count + v end
do local __t = tostring(count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
