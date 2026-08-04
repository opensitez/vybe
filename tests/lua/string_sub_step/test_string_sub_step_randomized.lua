-- vybe-test: lua/string_sub_step/test_string_sub_step_randomized
-- origin: languages/lua/tests/lua/test_string_sub_step.rs

local __w1 = "true"
local __i = 0

local t = {}
for i = 1, 24 do t[i] = string.sub("abcdefghijklmnopqrstuvwxyz", i, i) end
do local __t = tostring(#t == 24); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
