-- vybe-test: lua/control_flow_exhaustive/ctrl_if_elseif_branch
-- origin: languages/lua/tests/lua/test_control_flow_exhaustive.rs

local __w1 = "2"
local __i = 0

local x = 0; if false then x = 1 elseif true then x = 2 else x = 3 end; do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
