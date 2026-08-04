-- vybe-test: lua/string_gmatch/gmatch_digits
-- origin: languages/lua/tests/lua/test_string_gmatch.rs

local __w1 = "123,456,"
local __i = 0

local s=""
for d in string.gmatch("abc123def456", "%d+") do s=s..d.."," end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
