-- vybe-test: lua/programs/while_read_until_sentinel_pattern
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "3"
local __i = 0

local data = {1, 2, -1, 9}
local sum = 0
local i = 1
while data[i] ~= -1 do sum = sum + data[i] i = i + 1 end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
