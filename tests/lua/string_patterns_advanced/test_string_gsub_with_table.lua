-- vybe-test: lua/string_patterns_advanced/test_string_gsub_with_table
-- origin: languages/lua/tests/lua/test_string_patterns_advanced.rs

local __w1 = "A B c"
local __i = 0

local t={a='A', b='B'} do local __t = tostring((string.gsub('a b c', '%w', t))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
