-- vybe-test: lua/string_gmatch/gmatch_multiple_captures
-- origin: languages/lua/tests/lua/test_string_gmatch.rs

local __w1 = "x1,y2"
local __i = 0

local r={}
for a,b in string.gmatch("x:1 y:2", "(%a):(%d)") do r[#r+1]=a..b end
do local __t = tostring(table.concat(r,",")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
