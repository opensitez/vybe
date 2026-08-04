-- vybe-test: lua/numeric_modulo_sign/test_numeric_modulo_sign_chain_inner
-- origin: languages/lua/tests/lua/test_numeric_modulo_sign.rs

local __w1 = "0"
local __i = 0

do local __t = tostring((20 % 7) % 3); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
