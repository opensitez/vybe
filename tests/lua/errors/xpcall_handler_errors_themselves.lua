-- vybe-test: lua/errors/xpcall_handler_errors_themselves
-- origin: languages/lua/tests/lua/test_errors.rs

local __w1 = "false,false"
local __i = 0

local ok, msg = xpcall(function() error("first") end, function(e) error("second") end)
do local __t = tostring(tostring(ok) .. "," .. tostring(msg:match("second") ~= nil)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
