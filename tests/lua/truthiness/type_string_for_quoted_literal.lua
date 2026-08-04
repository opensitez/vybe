-- vybe-test: lua/truthiness/type_string_for_quoted_literal
-- origin: languages/lua/tests/lua/test_truthiness.rs

local __w1 = "string"
local __i = 0

do local __t = tostring(type("")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
