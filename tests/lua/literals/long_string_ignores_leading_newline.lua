-- vybe-test: lua/literals/long_string_ignores_leading_newline
-- origin: languages/lua/tests/lua/test_literals.rs

local __w1 = "hello"
local __i = 0

local s = [[
hello]]
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
