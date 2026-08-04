-- vybe-test: lua/errors/assert_false_triggers_error_message
-- origin: languages/lua/tests/lua/test_errors.rs

local __w1 = "bad"
local __i = 0

local ok, msg = pcall(function() assert(false, "bad") end)
do local __t = tostring(msg); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
