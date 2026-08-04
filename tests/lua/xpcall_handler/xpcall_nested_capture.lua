-- vybe-test: lua/xpcall_handler/xpcall_nested_capture
-- origin: languages/lua/tests/lua/test_xpcall_handler.rs

local __w1 = "false\ttrue"
local __i = 0

local inner_ok
local outer_ok = xpcall(function()
  inner_ok = xpcall(function() error("in") end, function() return "h" end)
end, function(e) return e end)
do local __t = tostring(inner_ok) .. "\t" .. tostring(outer_ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
