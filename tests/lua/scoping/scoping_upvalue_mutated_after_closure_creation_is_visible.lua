-- vybe-test: lua/scoping/scoping_upvalue_mutated_after_closure_creation_is_visible
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "20"
local __i = 0

local x = 10
local function get_x() return x end
x = 20
do local __t = tostring(get_x()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
