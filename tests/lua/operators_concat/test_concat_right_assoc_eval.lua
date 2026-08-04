-- vybe-test: lua/operators_concat/test_concat_right_assoc_eval
-- origin: languages/lua/tests/lua/test_operators_concat.rs

local __w1 = "3"
local __i = 0

local c=0; local function f(n) c=c+1 return n end; local _ = f(1) .. f(2) .. f(3); do local __t = tostring(c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
