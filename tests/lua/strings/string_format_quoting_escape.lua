-- vybe-test: lua/strings/string_format_quoting_escape
-- origin: languages/lua/tests/lua/test_strings.rs

local __w1 = "\"a\\n\""
local __i = 0

do local __t = tostring(string.format("%q", "a\n")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
