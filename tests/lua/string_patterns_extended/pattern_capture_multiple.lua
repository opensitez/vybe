-- vybe-test: lua/string_patterns_extended/pattern_capture_multiple
-- origin: languages/lua/tests/lua/test_string_patterns_extended.rs

local __w1 = "a:10"
local __i = 0

local x, y = string.match("a=10", "(%a+)=(%d+)")
do local __t = tostring(x .. ":" .. y); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
