-- vybe-test: lua/programs/swap_first_and_last_elements
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "30,10"
local __i = 0

local t = {10, 20, 30}
t[1], t[#t] = t[#t], t[1]
do local __t = tostring(t[1] .. "," .. t[#t]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
