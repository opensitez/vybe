-- vybe-test: lua/iteration/generic_for_with_manual_iterator
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "3"
local __i = 0

local function iter(_,i) i=i+1 if i>2 then return nil end return i,i end
local s=0
for _,v in iter,nil,0 do s=s+v end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
