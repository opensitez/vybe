-- vybe-test: lua/string_gsub_simple/test_string_gsub_simple_edge_second
-- origin: languages/lua/tests/lua/test_string_gsub_simple.rs

local __w1 = "true"
local __i = 0

local _, n = string.gsub("aaaaaaaaaaaaaaaaa", "a", "b")
do local __t = tostring(n == 17); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
