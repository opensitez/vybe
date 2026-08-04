-- vybe-test: lua/string_gsub_replacements/gsub_count_ret
-- origin: languages/lua/tests/lua/test_string_gsub_replacements.rs

local __w1 = "3"
local __i = 0

local _, n = string.gsub("banana", "a", "")
do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
