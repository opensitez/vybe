-- vybe-test: lua/lexical_scoping_advanced/lexical_param_shadow
-- origin: languages/lua/tests/lua/test_lexical_scoping_advanced.rs

local __w1 = "99"
local __i = 0

local x = 1
local function f(x)
  return function() return x end
end
do local __t = tostring(f(99)()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
