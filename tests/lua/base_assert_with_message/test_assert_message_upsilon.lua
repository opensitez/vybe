-- vybe-test: lua/base_assert_with_message/test_assert_message_upsilon
-- origin: languages/lua/tests/lua/test_base_assert_with_message.rs

local __w1 = "true"
local __i = 0

local ok, err = pcall(function() assert(false, "upsilon") end); do local __t = tostring(ok == false and string.find(tostring(err), "upsilon") ~= nil); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
