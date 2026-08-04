-- vybe-test: lua/string_gmatch/gmatch_chars
-- origin: languages/lua/tests/lua/test_string_gmatch.rs

local __w1 = "l-u-a-"
local __i = 0

local r=""
for c in string.gmatch("lua", ".") do r=r..c.."-" end
do local __t = tostring(r); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
