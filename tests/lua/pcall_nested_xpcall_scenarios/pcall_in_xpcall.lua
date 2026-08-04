-- vybe-test: lua/pcall_nested_xpcall_scenarios/pcall_in_xpcall
-- origin: languages/lua/tests/lua/test_pcall_nested_xpcall_scenarios.rs

local __w1 = "true\tfalse\tinput:5: inner"
local __i = 0

local function handler(e) return "h:" .. e end
local ok, val = xpcall(function()
  local ok2, res = pcall(function() error("inner") end)
  return ok2, res
end, handler)
do local __t = tostring(ok) .. "\t" .. tostring(val); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
