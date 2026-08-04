-- vybe-test: lua/functions/partial_application_via_closure
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "8"
local __i = 0

local function add(a, b) return a + b end
local function bind_a(a)
  return function(b) return add(a, b) end
end
do local __t = tostring(bind_a(5)(3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
