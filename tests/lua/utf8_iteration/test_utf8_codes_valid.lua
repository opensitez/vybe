-- vybe-test: lua/utf8_iteration/test_utf8_codes_valid
-- origin: languages/lua/tests/lua/test_utf8_iteration.rs

local __w1 = "1:65"
local __i = 0

local s=''; for p, c in utf8.codes('A') do s=s..p..':'..c end; do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
