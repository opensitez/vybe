-- vybe-test: lua/scope_variable_lifetime/local_for_loop_expire
-- origin: languages/lua/tests/lua/test_scope_variable_lifetime.rs

local __w1 = "nil"
local __i = 0

for i = 1, 3 do end
do local __t = tostring(tostring(i)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
