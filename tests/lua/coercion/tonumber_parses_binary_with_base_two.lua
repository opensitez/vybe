-- vybe-test: lua/coercion/tonumber_parses_binary_with_base_two
-- origin: languages/lua/tests/lua/test_coercion.rs

local __w1 = "10"
local __i = 0

do local __t = tostring(tonumber("1010", 2)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
