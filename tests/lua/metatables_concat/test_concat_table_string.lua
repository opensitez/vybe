-- vybe-test: lua/metatables_concat/test_concat_table_string
-- origin: languages/lua/tests/lua/test_metatables_concat.rs

local __w1 = "ab"
local __i = 0

local mt={__concat=function(a,b) return a.v..b end}; local t1=setmetatable({v='a'}, mt); do local __t = tostring(t1..'b'); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
