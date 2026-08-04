-- vybe-test: lua/utf8_validation/test_utf8_offset_out_of_bounds
-- origin: languages/lua/tests/lua/test_utf8_validation.rs

local __w1 = "nil"
local __i = 0

do local __t = tostring(utf8.offset('a', 3) or 'nil'); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
