-- vybe-test: lua/scoping_exhaustive/scope_exh_global_shadowed
-- origin: languages/lua/tests/lua/test_scoping_exhaustive.rs

local __w1 = "200\t100"
local __i = 0

g_val = 100
local g_val = 200
do local __t = tostring(g_val) .. "\t" .. tostring(_G.g_val); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
