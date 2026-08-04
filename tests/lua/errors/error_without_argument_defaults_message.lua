-- vybe-test: lua/errors/error_without_argument_defaults_message
-- origin: languages/lua/tests/lua/test_errors.rs

local __w1 = "string"
local __i = 0

local ok, msg = pcall(function() error() end)
do local __t = tostring(type(msg)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
