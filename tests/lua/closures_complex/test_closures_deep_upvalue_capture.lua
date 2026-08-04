-- vybe-test: lua/closures_complex/test_closures_deep_upvalue_capture
-- origin: languages/lua/tests/lua/test_closures_complex.rs

local __w1 = "60"
local __i = 0

local function outer(x)
    local function inner(y)
        local function deepest(z)
            return x + y + z
        end
        return deepest
    end
    return inner
end
local f = outer(10)
local g = f(20)
do local __t = tostring(g(30)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
