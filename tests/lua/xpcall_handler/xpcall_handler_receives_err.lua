-- vybe-test: lua/xpcall_handler/xpcall_handler_receives_err
-- origin: languages/lua/tests/lua/test_xpcall_handler.rs

local __w1 = "false\tcaught:input:1: boom"
local __i = 0

local function handler(err) return "caught:" .. tostring(err) end
local ok, msg = xpcall(function() error("boom") end, handler)
do local __t = tostring(ok) .. "\t" .. tostring(msg); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
