-- vybe-test: lua/tables_mixed/test_mixed_table_ipairs_stops_at_hole
-- origin: languages/lua/tests/lua/test_tables_mixed.rs

local __w1 = "1"
local __i = 0

local t={1, nil, 3, a=4}; local c=0; for k,v in ipairs(t) do c=c+1 end; do local __t = tostring(c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
