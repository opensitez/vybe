-- vybe-test: lua/scoping_exhaustive/scope_exh_upvalue_nested
-- origin: languages/lua/tests/lua/test_scoping_exhaustive.rs

local __w1 = "42"
local __i = 0

local x = 42
local function f() return function() return x end end
do local __t = tostring(f()()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
