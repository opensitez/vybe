-- vybe-test: lua/string_find/test_find_not_found
-- origin: languages/lua/tests/lua/test_string_find.rs

local __w1 = "nil"
local __i = 0

local s = string.find('hello world', 'lua'); do local __t = tostring(tostring(s)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
