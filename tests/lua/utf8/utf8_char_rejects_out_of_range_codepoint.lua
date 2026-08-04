-- vybe-test: lua/utf8/utf8_char_rejects_out_of_range_codepoint
-- origin: languages/lua/tests/lua/test_utf8.rs

local __w1 = "false"
local __i = 0

local ok = pcall(function() utf8.char(0x110000) end)
do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
