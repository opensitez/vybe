-- vybe-test: lua/string_gmatch/gmatch_key_val
-- origin: languages/lua/tests/lua/test_string_gmatch.rs

local __w1 = "1,2"
local __i = 0

local t={}
for k,v in string.gmatch("a=1,b=2", "(%a+)=(%d+)") do t[k]=v end
do local __t = tostring(t["a"] .. "," .. t["b"]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
