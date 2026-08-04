-- vybe-test: lua/strings/string_gsub_function_retains_nil_results
-- origin: languages/lua/tests/lua/test_strings.rs

local __w1 = "a B c"
local __i = 0

local res = string.gsub("a b c", "%a", function(x) if x == "b" then return "B" end end)
do local __t = tostring(res); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
