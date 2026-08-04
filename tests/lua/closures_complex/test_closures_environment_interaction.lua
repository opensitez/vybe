-- vybe-test: lua/closures_complex/test_closures_environment_interaction
-- origin: languages/lua/tests/lua/test_closures_complex.rs

local __w1 = "1"
local __i = 0

local x = 1
local function get_x() return x end
local _ENV = {x = 10, get_x = get_x}
do local __t = tostring(get_x()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
