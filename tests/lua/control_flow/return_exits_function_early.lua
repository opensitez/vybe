-- vybe-test: lua/control_flow/return_exits_function_early
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "5"
local __i = 0

function f()
  if true then return 5 end
  return 0
end
do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
