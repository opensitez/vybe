-- vybe-test: lua/metatables_call/test_call_table
-- origin: languages/lua/tests/lua/test_metatables_call.rs

local __w1 = "30"
local __i = 0

local t=setmetatable({}, {__call=function(tbl, a, b) return a+b end}); do local __t = tostring(t(10, 20)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
