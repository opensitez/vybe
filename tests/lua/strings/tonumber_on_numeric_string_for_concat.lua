-- vybe-test: lua/strings/tonumber_on_numeric_string_for_concat
-- origin: languages/lua/tests/lua/test_strings.rs

local __w1 = "n=42"
local __i = 0

local n = tonumber("42")
do local __t = tostring("n=" .. n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
