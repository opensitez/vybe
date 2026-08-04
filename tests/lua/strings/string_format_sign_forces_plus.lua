-- vybe-test: lua/strings/string_format_sign_forces_plus
-- origin: languages/lua/tests/lua/test_strings.rs

local __w1 = "+42 -42"
local __i = 0

do local __t = tostring(string.format("%+d", 42) .. " " .. string.format("%+d", -42)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
