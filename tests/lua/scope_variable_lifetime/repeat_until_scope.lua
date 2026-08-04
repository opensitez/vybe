-- vybe-test: lua/scope_variable_lifetime/repeat_until_scope
-- origin: languages/lua/tests/lua/test_scope_variable_lifetime.rs

local __w1 = "true"
local __i = 0

local done = false
repeat
  local x = 1
done = (x == 1)
until done
do local __t = tostring(done); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
