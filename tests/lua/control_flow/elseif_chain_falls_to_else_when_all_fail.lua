-- vybe-test: lua/control_flow/elseif_chain_falls_to_else_when_all_fail
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "none"
local __i = 0

local x = 5
if x == 1 then do local __t = tostring('a'); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
elseif x == 2 then do local __t = tostring('b'); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
elseif x == 3 then do local __t = tostring('c'); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
else do local __t = tostring('none'); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
