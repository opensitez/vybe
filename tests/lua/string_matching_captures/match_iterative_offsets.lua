-- vybe-test: lua/string_matching_captures/match_iterative_offsets
-- origin: languages/lua/tests/lua/test_string_matching_captures.rs

local __w1 = "one-two-three"
local __i = 0

local s = "one two three"
local words = {}
for w in string.gmatch(s, "%a+") do
  words[#words+1] = w
end
do local __t = tostring(table.concat(words, "-")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
