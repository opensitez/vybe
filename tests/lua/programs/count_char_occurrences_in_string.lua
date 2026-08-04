-- vybe-test: lua/programs/count_char_occurrences_in_string
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "3"
local __i = 0

local s, ch, n = "banana", "a", 0
for i = 1, #s do if string.sub(s, i, i) == ch then n = n + 1 end end
do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
