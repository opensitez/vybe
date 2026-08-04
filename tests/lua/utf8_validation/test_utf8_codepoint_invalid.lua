-- vybe-test: lua/utf8_validation/test_utf8_codepoint_invalid
-- origin: languages/lua/tests/lua/test_utf8_validation.rs

local __w1 = "false"
local __i = 0

local ok, err = pcall(function() utf8.codepoint('a\xFFb', 1, 3) end); do local __t = tostring(tostring(ok)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
