-- vybe-test: lua/tables/table_reference_keys
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "one,two"
local __i = 0

local t = {}
local k1 = {}
local k2 = {}
t[k1] = "one"
t[k2] = "two"
do local __t = tostring(t[k1] .. "," .. t[k2]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
