-- vybe-test: lua/control_flow/guard_clause_style_early_exit
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "nil"
local __i = 0

local function f(x)
  if x == nil then return "nil" end
  return "ok"
end
do local __t = tostring(f(nil)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
