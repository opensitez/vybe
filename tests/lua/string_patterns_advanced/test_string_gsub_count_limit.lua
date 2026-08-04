-- vybe-test: lua/string_patterns_advanced/test_string_gsub_count_limit
-- origin: languages/lua/tests/lua/test_string_patterns_advanced.rs

local __w1 = "b b b a a 3"
local __i = 0

local res, cnt = string.gsub('a a a a a', 'a', 'b', 3) do local __t = tostring(res .. ' ' .. cnt); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
