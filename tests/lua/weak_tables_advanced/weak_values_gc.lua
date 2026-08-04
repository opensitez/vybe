-- vybe-test: lua/weak_tables_advanced/weak_values_gc
-- origin: languages/lua/tests/lua/test_weak_tables_advanced.rs

local __w1 = "true"
local __i = 0

local t = setmetatable({}, {__mode="v"})
local key = {}
local val = {}
t[key] = val
do local __t = tostring(t[key] == val); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
