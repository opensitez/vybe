-- vybe-test: lua/string_pack_overflow/test_string_pack_overflow_unicode_like
-- origin: languages/lua/tests/lua/test_string_pack_overflow.rs

local __w1 = "true"
local __i = 0

local ok = pcall(function() string.pack("b", 147) end)
do local __t = tostring(type(ok) == "boolean"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
