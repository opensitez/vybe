-- vybe-test: lua/metatable_protection/metatable_protected_val_table
-- origin: languages/lua/tests/lua/test_metatable_protection.rs

local __w1 = "true"
local __i = 0

local guard = {}
local mt = {__metatable = guard}
local t = setmetatable({}, mt)
do local __t = tostring(getmetatable(t) == guard); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
