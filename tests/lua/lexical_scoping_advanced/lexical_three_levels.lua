-- vybe-test: lua/lexical_scoping_advanced/lexical_three_levels
-- origin: languages/lua/tests/lua/test_lexical_scoping_advanced.rs

local __w1 = "60"
local __i = 0

local x = 10
local function f1()
  local y = 20
  return function()
    local z = 30
    return x + y + z
  end
end
do local __t = tostring(f1()()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
