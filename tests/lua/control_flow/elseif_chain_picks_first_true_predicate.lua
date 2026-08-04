-- vybe-test: lua/control_flow/elseif_chain_picks_first_true_predicate
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "pos"
local __i = 0

local n = 5
if n < 0 then do local __t = tostring("neg"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
elseif n == 0 then do local __t = tostring("zero"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
elseif n > 0 then do local __t = tostring("pos"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
