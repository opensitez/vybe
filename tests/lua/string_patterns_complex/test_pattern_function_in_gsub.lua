-- vybe-test: lua/string_patterns_complex/test_pattern_function_in_gsub
-- origin: languages/lua/tests/lua/test_string_patterns_complex.rs

local __w1 = "A B C"
local __i = 0

local s = 'a b c'
local r = string.gsub(s, '%w+', function(w) return string.upper(w) end)
do local __t = tostring(r); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
