-- vybe-test: lua/iteration/numeric_for_zero_step_is_invalid_at_runtime
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "false"
local __i = 0

local ok,err=pcall(function() for i=1,3,0 do end end)
do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
