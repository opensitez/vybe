-- vybe-test: lua/strings/count_char_occurrence_manual_loop
-- origin: languages/lua/tests/lua/test_strings.rs

local __w1 = "3"
local __i = 0

local s = "banana"
local n = 0
for i = 1, #s do if string.sub(s,i,i) == "a" then n = n + 1 end end
do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
