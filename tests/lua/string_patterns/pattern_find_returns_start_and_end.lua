-- vybe-test: lua/string_patterns/pattern_find_returns_start_and_end
-- origin: languages/lua/tests/lua/test_string_patterns.rs

local __w1 = "2,3"
local __i = 0

local s,e=string.find("banana", "an")
do local __t = tostring(s..","..e); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
