-- vybe-test: lua/environment/rawget_on_g_for_global_lookup
-- origin: languages/lua/tests/lua/test_environment.rs

local __w1 = "42"
local __i = 0

answer = 42
do local __t = tostring(rawget(_G, 'answer')); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
