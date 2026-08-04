-- vybe-test: lua/metatables_call/test_call_chain
-- origin: languages/lua/tests/lua/test_metatables_call.rs

local __w1 = "10"
local __i = 0

local t1=setmetatable({}, {__call=function(tbl, a) return a*2 end}); local t2=setmetatable({}, {__call=t1}); do local __t = tostring(t2(5)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
