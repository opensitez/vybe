-- vybe-test: lua/error_handling_xpcall/test_xpcall_error_in_handler
-- origin: languages/lua/tests/lua/test_error_handling_xpcall.rs

local __w1 = "false true"
local __i = 0

local ok, err = xpcall(function() error('boom1') end, function() error('boom2') end); do local __t = tostring(tostring(ok)..' '..tostring(string.find(err, 'error in error handling') ~= nil)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
