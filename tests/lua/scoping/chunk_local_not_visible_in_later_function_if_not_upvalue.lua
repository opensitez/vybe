-- vybe-test: lua/scoping/chunk_local_not_visible_in_later_function_if_not_upvalue
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "6"
local __i = 0

local z = 5
function g() return z end
z = 6
do local __t = tostring(g()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
