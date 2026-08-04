-- vybe-test: lua/scoping_exhaustive/scope_exh_global_fallback
-- origin: languages/lua/tests/lua/test_scoping_exhaustive.rs

local __w1 = "99"
local __i = 0

g_val = 99
do local __t = tostring(g_val); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
