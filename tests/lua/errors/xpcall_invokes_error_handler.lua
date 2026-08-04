-- vybe-test: lua/errors/xpcall_invokes_error_handler
-- origin: languages/lua/tests/lua/test_errors.rs

local __w1 = "handled:boom"
local __i = 0

local function handler(err) return "handled:" .. err end
local ok, msg = xpcall(function() error("boom") end, handler)
do local __t = tostring(msg); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
