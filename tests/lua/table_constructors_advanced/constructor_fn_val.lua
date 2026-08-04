-- vybe-test: lua/table_constructors_advanced/constructor_fn_val
-- origin: languages/lua/tests/lua/test_table_constructors_advanced.rs

local __w1 = "10"
local __i = 0

local t = {fn=function(x) return x*2 end}
do local __t = tostring(t.fn(5)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
