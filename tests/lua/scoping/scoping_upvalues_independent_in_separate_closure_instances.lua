-- vybe-test: lua/scoping/scoping_upvalues_independent_in_separate_closure_instances
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "3 1"
local __i = 0

local function make_counter()
  local count = 0
  return function() count = count + 1; return count end
end
local c1 = make_counter()
local c2 = make_counter()
c1(); c1()
do local __t = tostring(c1() .. " " .. c2()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
