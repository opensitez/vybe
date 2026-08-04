-- vybe-test: lua/tables_metatables_ext/test_metatable_protect_get
-- origin: languages/lua/tests/lua/test_tables_metatables_ext.rs

local __w1 = "protected"
local __i = 0

local t={}; setmetatable(t, {__metatable='protected'}); do local __t = tostring(getmetatable(t)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
