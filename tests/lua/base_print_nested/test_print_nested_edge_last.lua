-- vybe-test: lua/base_print_nested/test_print_nested_edge_last
-- origin: languages/lua/tests/lua/test_base_print_nested.rs

local __w1 = "37"
local __i = 0

do local __t = tostring((function() return (function(x) return (function(y) return x + y end)(18) end)(19) end)()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
