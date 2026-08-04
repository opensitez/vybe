-- vybe-test: lua/iteration/numeric_for_accumulates_string
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "1234"
local __i = 0

local s = ""
for i = 1, 4 do s = s .. i end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
