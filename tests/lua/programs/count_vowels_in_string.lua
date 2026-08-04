-- vybe-test: lua/programs/count_vowels_in_string
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "2"
local __i = 0

local s = "hello"
local n = 0
for i = 1, #s do
  local c = string.sub(s, i, i)
  if c == "a" or c == "e" or c == "i" or c == "o" or c == "u" then n = n + 1 end
end
do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
