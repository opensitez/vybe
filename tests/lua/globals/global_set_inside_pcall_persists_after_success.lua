-- vybe-test: lua/globals/global_set_inside_pcall_persists_after_success
-- origin: languages/lua/tests/lua/test_globals.rs

local __w1 = "set"
local __i = 0

pcall(function() side_effect_global = 'set' end)
do local __t = tostring(side_effect_global); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
