-- vybe-test: lua/operators/length_operator_on_string_gives_byte_count
-- origin: languages/lua/tests/lua/test_operators.rs

local __w1 = "5"
local __i = 0

local s = 'hello'
do local __t = tostring(#s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
