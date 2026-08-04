-- vybe-test: lua/errors/xpcall_returns_false_on_failure
-- origin: languages/lua/tests/lua/test_errors.rs

local __w1 = "false"
local __i = 0

local ok = xpcall(function() error("e") end, function(e) return e end)
do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
