-- vybe-test: lua/xpcall_handler/xpcall_raw_err
-- origin: languages/lua/tests/lua/test_xpcall_handler.rs

local __w1 = "raw"
local __i = 0

local handler_got = nil
xpcall(function() error("raw", 0) end, function(e) handler_got = e end)
do local __t = tostring(handler_got); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
