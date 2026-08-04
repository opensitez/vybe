-- vybe-test: lua/programs/reverse_words_in_string
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "three two one"
local __i = 0

local s = "one two three"
local rev = ""
for word in string.gmatch(s, "%S+") do
  rev = word .. (rev == "" and "" or " ") .. rev
end
do local __t = tostring(rev); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
