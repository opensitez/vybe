-- vybe-test: lua/coercion/tonumber_stops_at_invalid_suffix
-- origin: languages/lua/tests/lua/test_coercion.rs

local __w1 = "nil"
local __i = 0

do local __t = tostring(tostring(tonumber("42px"))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
