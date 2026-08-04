-- vybe-test: lua/control_flow/elseif_chain
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "2"
local __i = 0

if false then do local __t = tostring(1); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
elseif true then do local __t = tostring(2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
else do local __t = tostring(3); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
