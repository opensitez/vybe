-- vybe-test: lua/string_patterns_complex/test_pattern_table_in_gsub
-- origin: languages/lua/tests/lua/test_string_patterns_complex.rs

local __w1 = "alpha and beta"
local __i = 0

local t = {a = 'alpha', b = 'beta'}
local r = string.gsub('a and b', '%w+', t)
do local __t = tostring(r); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
