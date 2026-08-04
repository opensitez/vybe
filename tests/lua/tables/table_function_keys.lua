-- vybe-test: lua/tables/table_function_keys
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "300"
local __i = 0

local t = {}
local f1 = function() end
local f2 = function() end
t[f1] = 100
t[f2] = 200
do local __t = tostring(t[f1] + t[f2]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
