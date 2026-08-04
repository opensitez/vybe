-- vybe-test: lua/types_exhaustive/str_escape_quote
-- origin: languages/lua/tests/lua/test_types_exhaustive.rs

local __w1 = "a\"b"
local __i = 0

do local __t = tostring("a\"b"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
