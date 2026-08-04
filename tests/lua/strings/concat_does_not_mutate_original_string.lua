-- vybe-test: lua/strings/concat_does_not_mutate_original_string
-- origin: languages/lua/tests/lua/test_strings.rs

local __w1 = "ab,abc"
local __i = 0

local s = "ab"
local t = s .. "c"
do local __t = tostring(s .. "," .. t); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
