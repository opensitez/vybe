-- vybe-test: lua/error_handling_xpcall/test_xpcall_error_handler_return_value
-- origin: languages/lua/tests/lua/test_error_handling_xpcall.rs

local __w1 = "false my_error"
local __i = 0

local ok, err = xpcall(function() error('boom') end, function(e) return 'my_error' end); do local __t = tostring(tostring(ok)..' '..err); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
