-- vybe-test: lua/errors/pcall_with_multiple_return_values
-- origin: languages/lua/tests/lua/test_errors.rs

local __w1 = "3"
local __i = 0

local ok, a, b = pcall(function() return 1, 2 end)
do local __t = tostring(a + b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
