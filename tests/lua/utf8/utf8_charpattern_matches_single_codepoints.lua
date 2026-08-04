-- vybe-test: lua/utf8/utf8_charpattern_matches_single_codepoints
-- origin: languages/lua/tests/lua/test_utf8.rs

local __w1 = "3"
local __i = 0

local count = 0
for _ in string.gmatch("aλb", utf8.charpattern) do count = count + 1 end
do local __t = tostring(count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
