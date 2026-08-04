-- vybe-test: lua/closures/global_lookup_when_not_local
-- origin: languages/lua/tests/lua/test_closures.rs

local __w1 = "5"
local __i = 0

gval=5
local function f() return gval end
do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
