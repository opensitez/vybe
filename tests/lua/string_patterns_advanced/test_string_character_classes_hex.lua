-- vybe-test: lua/string_patterns_advanced/test_string_character_classes_hex
-- origin: languages/lua/tests/lua/test_string_patterns_advanced.rs

local __w1 = "1A2F"
local __i = 0

do local __t = tostring((string.gsub('g1hA2zF', '%X', ''))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
