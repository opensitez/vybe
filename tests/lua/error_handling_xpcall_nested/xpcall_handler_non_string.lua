-- vybe-test: lua/error_handling_xpcall_nested/xpcall_handler_non_string
-- origin: languages/lua/tests/lua/test_error_handling_xpcall_nested.rs

local __w1 = "false table 500"
local __i = 0

local ok, err = xpcall(function() error("boom") end, function() return {code=500} end)
do local __t = tostring(ok) .. "\t" .. tostring(type(err)) .. "\t" .. tostring(err.code); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
