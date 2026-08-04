-- vybe-test: lua/string_char_byte/test_char_invalid_arg
-- origin: languages/lua/tests/lua/test_string_char_byte.rs

local __w1 = "false"
local __i = 0

local ok = pcall(function() string.char(300) end); do local __t = tostring(tostring(ok)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
