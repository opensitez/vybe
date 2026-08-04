-- vybe-test: lua/xpcall_handler/xpcall_no_error_no_handler
-- origin: languages/lua/tests/lua/test_xpcall_handler.rs

local __w1 = "true\tfalse"
local __i = 0

local called = false
local ok = xpcall(function() return 1 end, function() called = true end)
do local __t = tostring(ok) .. "\t" .. tostring(called); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
