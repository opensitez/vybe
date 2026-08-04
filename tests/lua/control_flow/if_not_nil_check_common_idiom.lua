-- vybe-test: lua/control_flow/if_not_nil_check_common_idiom
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "set"
local __i = 0

local v = 1
if v ~= nil then do local __t = tostring("set"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
