-- vybe-test: lua/string_patterns_advanced/test_string_character_classes_digits
-- origin: languages/lua/tests/lua/test_string_patterns_advanced.rs

local __w1 = "1234"
local __i = 0

do local __t = tostring((string.gsub('a1b2c3D4', '%D', ''))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
