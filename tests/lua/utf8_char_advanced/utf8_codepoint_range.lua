-- vybe-test: lua/utf8_char_advanced/utf8_codepoint_range
-- origin: languages/lua/tests/lua/test_utf8_char_advanced.rs

local __w1 = "65,945"
local __i = 0

local s = "Aα"
local a, b = utf8.codepoint(s, 1, #s)
do local __t = tostring(a .. "," .. b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
