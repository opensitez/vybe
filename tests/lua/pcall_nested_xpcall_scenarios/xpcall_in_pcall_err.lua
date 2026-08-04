-- vybe-test: lua/pcall_nested_xpcall_scenarios/xpcall_in_pcall_err
-- origin: languages/lua/tests/lua/test_pcall_nested_xpcall_scenarios.rs

local __w1 = "true\tfalse\thandled:fail"
local __i = 0

local ok, inner_ok, val = pcall(function()
  return xpcall(function() error("fail", 0) end, function(e) return "handled:"..e end)
end)
do local __t = tostring(ok) .. "\t" .. tostring(inner_ok) .. "\t" .. tostring(val); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
