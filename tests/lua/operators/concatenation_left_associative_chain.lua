-- vybe-test: lua/operators/concatenation_left_associative_chain
-- origin: languages/lua/tests/lua/test_operators.rs

local __w1 = "abcd"
local __i = 0

do local __t = tostring("a".."b".."c".."d"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
