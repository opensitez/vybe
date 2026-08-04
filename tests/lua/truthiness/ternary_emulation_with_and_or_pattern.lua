-- vybe-test: lua/truthiness/ternary_emulation_with_and_or_pattern
-- origin: languages/lua/tests/lua/test_truthiness.rs

local __w1 = "yes"
local __i = 0

local x = true
local result = x and 'yes' or 'no'
do local __t = tostring(result); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
