-- vybe-test: lua/closures_complex/test_closures_escaping_upvalues
-- origin: languages/lua/tests/lua/test_closures_complex.rs

local __w1 = "42"
local __i = 0

local f
do
    local x = 42
    f = function() return x end
end
do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
