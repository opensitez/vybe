-- vybe-test: lua/error_handling_complex/test_xpcall_message_handler_modifies_error
-- origin: languages/lua/tests/lua/test_error_handling_complex.rs

local __w1 = "caught: boom"
local __i = 0

local function handler(err)
    return 'caught: ' .. tostring(err)
end
local ok, res = xpcall(function() error('boom') end, handler)
do local __t = tostring(res); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
