-- vybe-test: lua/control_flow_exhaustive/ctrl_for_gen_ipairs
-- origin: languages/lua/tests/lua/test_control_flow_exhaustive.rs

local __w1 = "ab"
local __i = 0

local s = ""
for i, v in ipairs({"a", "b"}) do s = s .. v end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
