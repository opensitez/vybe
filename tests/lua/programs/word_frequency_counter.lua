-- vybe-test: lua/programs/word_frequency_counter
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "3,2,1"
local __i = 0

local text = 'the cat sat on the mat the cat'
local freq = {}
for word in text:gmatch('%a+') do
  freq[word] = (freq[word] or 0) + 1
end
do local __t = tostring(freq['the'] .. ',' .. freq['cat'] .. ',' .. freq['sat']); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
