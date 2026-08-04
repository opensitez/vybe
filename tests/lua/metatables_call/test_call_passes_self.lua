-- vybe-test: lua/metatables_call/test_call_passes_self
-- origin: languages/lua/tests/lua/test_metatables_call.rs

local __w1 = "true"
local __i = 0

local target; local t=setmetatable({}, {__call=function(tbl) target=tbl end}); t(); do local __t = tostring(tostring(t==target)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
