-- vybe-test: lua/scoping_exhaustive/scope_exh_sibling_sharing
-- origin: languages/lua/tests/lua/test_scoping_exhaustive.rs

local __w1 = "1"
local __i = 0

local x = 0
local f1 = function() x = x + 1 end
local f2 = function() return x end
f1()
do local __t = tostring(f2()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
