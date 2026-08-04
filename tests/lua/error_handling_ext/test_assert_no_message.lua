-- vybe-test: lua/error_handling_ext/test_assert_no_message
-- origin: languages/lua/tests/lua/test_error_handling_ext.rs

local __w1 = "true"
local __i = 0

local ok, err = pcall(function() assert(false) end); do local __t = tostring(tostring(string.find(err, 'assertion failed') ~= nil)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
