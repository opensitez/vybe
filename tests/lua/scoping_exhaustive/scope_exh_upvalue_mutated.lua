-- vybe-test: lua/scoping_exhaustive/scope_exh_upvalue_mutated
-- origin: languages/lua/tests/lua/test_scoping_exhaustive.rs

local __w1 = "11"
local __i = 0

local x = 10
local function f() x = x + 1 end
f()
do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
