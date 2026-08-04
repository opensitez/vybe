-- vybe-test: lua/string_patterns_frontier_balanced/frontier_before_lower
-- origin: languages/lua/tests/lua/test_string_patterns_frontier_balanced.rs

local __w1 = "2"
local __i = 0

local s = "hello world"
local t = {}
for w in string.gmatch(s, "%f[%a]%a+") do t[#t+1] = w end
do local __t = tostring(#t); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
