-- vybe-test: lua/string_find/test_find_captures
-- origin: languages/lua/tests/lua/test_string_find.rs

local __w1 = "3 7 12 34"
local __i = 0

local s, e, c1, c2 = string.find('a 12 34 b', '(%d+) (%d+)'); do local __t = tostring(s..' '..e..' '..c1..' '..c2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
