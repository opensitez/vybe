-- vybe-test: lua/strings/pad_left_with_rep_and_concat
-- origin: languages/lua/tests/lua/test_strings.rs

local __w1 = "007"
local __i = 0

local s = "7"
do local __t = tostring(string.rep("0", 3 - #s) .. s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
