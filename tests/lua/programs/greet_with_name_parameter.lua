-- vybe-test: lua/programs/greet_with_name_parameter
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "hi ada"
local __i = 0

local function greet(name) return "hi " .. name end
do local __t = tostring(greet("ada")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
