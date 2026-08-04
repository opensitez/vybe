-- vybe-test: lua/scoping/function_parameter_shadows_outer_local_in_body
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "9"
local __i = 0

local n = 1
function f(n) return n end
do local __t = tostring(f(9)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
