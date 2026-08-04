-- vybe-test: lua/control_flow/if_elseif_else_with_local_result
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "pos"
local __i = 0

local n = 2
local label
if n < 0 then label = "neg"
elseif n == 0 then label = "zero"
else label = "pos" end
do local __t = tostring(label); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
