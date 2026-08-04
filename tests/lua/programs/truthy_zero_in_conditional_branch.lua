-- vybe-test: lua/programs/truthy_zero_in_conditional_branch
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "truthy"
local __i = 0

local n = 0
if n then do local __t = tostring("truthy"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end else do local __t = tostring("falsy"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
