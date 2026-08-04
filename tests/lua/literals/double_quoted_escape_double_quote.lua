-- vybe-test: lua/literals/double_quoted_escape_double_quote
-- origin: languages/lua/tests/lua/test_literals.rs

local __w1 = "\""
local __i = 0

do local __t = tostring("\""); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
