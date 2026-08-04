-- vybe-test: lua/programs/tokenize_words_with_pattern_gmatch
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "2"
local __i = 0

local s = "one two"
local n = 0
for _ in string.gmatch(s, "%a+") do n = n + 1 end
do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
