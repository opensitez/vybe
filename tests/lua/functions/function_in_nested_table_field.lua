-- vybe-test: lua/functions/function_in_nested_table_field
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "42"
local __i = 0

local api = { utils = { double = function(x) return x * 2 end } }
do local __t = tostring(api.utils.double(21)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
