-- vybe-test: lua/string_patterns/pattern_gmatch_yields_all_words
-- origin: languages/lua/tests/lua/test_string_patterns.rs

local __w1 = "one,two"
local __i = 0

local t={}
for w in string.gmatch("one two", "%S+") do t[#t+1]=w end
do local __t = tostring(table.concat(t,",")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
