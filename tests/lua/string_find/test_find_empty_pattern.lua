-- vybe-test: lua/string_find/test_find_empty_pattern
-- origin: languages/lua/tests/lua/test_string_find.rs

local __w1 = "1 0"
local __i = 0

local s, e = string.find('abc', '()'); do local __t = tostring(s..' '..e); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
