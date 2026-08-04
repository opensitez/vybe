-- vybe-test: lua/string_patterns_advanced/test_string_find_position_capture
-- origin: languages/lua/tests/lua/test_string_patterns_advanced.rs

local __w1 = "3 5"
local __i = 0

local s, e, p1, p2 = string.find('hello', '()ll()') do local __t = tostring(p1 .. ' ' .. p2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
