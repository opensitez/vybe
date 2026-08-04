-- vybe-test: lua/tables_sort_custom/test_sort_custom_stable
-- origin: languages/lua/tests/lua/test_tables_sort_custom.rs

local __w1 = "21"
local __i = 0

local t={{a=1,b=2},{a=1,b=1}}; table.sort(t, function(x,y) return x.a<y.a end); do local __t = tostring(t[1].b..t[2].b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
