-- vybe-test: lua/string_gmatch/gmatch_no_captures_returns_full_match
-- origin: languages/lua/tests/lua/test_string_gmatch.rs

local __w1 = "cat,bat,sat"
local __i = 0

local r={}
for m in string.gmatch("cat bat sat", "%a+at") do r[#r+1]=m end
do local __t = tostring(table.concat(r,",")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
