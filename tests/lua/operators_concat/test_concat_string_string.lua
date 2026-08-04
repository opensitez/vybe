-- vybe-test: lua/operators_concat/test_concat_string_string
-- origin: languages/lua/tests/lua/test_operators_concat.rs

local __w1 = "abcdef"
local __i = 0

do local __t = tostring('abc' .. 'def'); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
