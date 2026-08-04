-- vybe-test: lua/nan_and_inf/tonumber_parses_nan_string
-- origin: languages/lua/tests/lua/test_nan_and_inf.rs

local __w1 = "true"
local __i = 0

local n = tonumber("nan")
do local __t = tostring(n ~= n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
