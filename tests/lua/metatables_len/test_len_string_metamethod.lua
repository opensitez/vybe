-- vybe-test: lua/metatables_len/test_len_string_metamethod
-- origin: languages/lua/tests/lua/test_metatables_len.rs

local __w1 = "99"
local __i = 0

debug.setmetatable('', {__len=function(s) return 99 end}); do local __t = tostring(#'abc'); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
