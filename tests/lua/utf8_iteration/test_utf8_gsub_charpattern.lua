-- vybe-test: lua/utf8_iteration/test_utf8_gsub_charpattern
-- origin: languages/lua/tests/lua/test_utf8_iteration.rs

local __w1 = "XXX"
local __i = 0

local s = string.gsub('A你B', utf8.charpattern, 'X'); do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
