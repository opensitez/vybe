-- vybe-test: lua/scope_variable_lifetime/global_block_survival
-- origin: languages/lua/tests/lua/test_scope_variable_lifetime.rs

local __w1 = "42"
local __i = 0

do g_var = 42 end
do local __t = tostring(g_var); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
