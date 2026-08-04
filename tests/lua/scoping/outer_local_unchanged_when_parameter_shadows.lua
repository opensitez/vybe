-- vybe-test: lua/scoping/outer_local_unchanged_when_parameter_shadows
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "1"
local __i = 0

local n = 1
function f(n) end
do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
