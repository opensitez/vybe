-- vybe-test: lua/literals_long_strings/test_long_string_ignore_first_newline
-- origin: languages/lua/tests/lua/test_literals_long_strings.rs

local __w1 = "\\nhello"
local __i = 0

do local __t = tostring([[\nhello]]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
