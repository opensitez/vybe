-- vybe-test: lua/strings/string_gsub_table_ignores_missing_keys
-- origin: languages/lua/tests/lua/test_strings.rs

local __w1 = "A b c"
local __i = 0

local res = string.gsub("a b c", "%a", {a="A"})
do local __t = tostring(res); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
