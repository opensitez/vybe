-- vybe-test: lua/operators/string_concat_three_parts
-- origin: languages/lua/tests/lua/test_operators.rs

local __w1 = "abc"
local __i = 0

do local __t = tostring("a".."b".."c"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
