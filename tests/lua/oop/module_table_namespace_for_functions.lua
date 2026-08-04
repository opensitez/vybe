-- vybe-test: lua/oop/module_table_namespace_for_functions
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "7"
local __i = 0

local List = {}
function List.head(t) return t[1] end
do local __t = tostring(List.head({7, 8})); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
