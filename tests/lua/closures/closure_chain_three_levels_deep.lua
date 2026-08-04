-- vybe-test: lua/closures/closure_chain_three_levels_deep
-- origin: languages/lua/tests/lua/test_closures.rs

local __w1 = "6"
local __i = 0

local function a(x)
  return function(y)
    return function(z) return x + y + z end
  end
end
do local __t = tostring(a(1)(2)(3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
