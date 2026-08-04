-- vybe-test: lua/type_checks/next_invalid_key_raises_error
-- origin: languages/lua/tests/lua/test_type_checks.rs

local __w1 = "false"
local __i = 0

local ok, err = pcall(function() next({a=1}, "invalid_key") end)
do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
