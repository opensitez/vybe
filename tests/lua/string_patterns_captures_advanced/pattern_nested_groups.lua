-- vybe-test: lua/string_patterns_captures_advanced/pattern_nested_groups
-- origin: languages/lua/tests/lua/test_string_patterns_captures_advanced.rs

local __w1 = "hello(world),world"
local __i = 0

local outer, inner = string.match("hello(world)", "(%a+%((%a+)%))")
do local __t = tostring(outer .. "," .. inner); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
