-- vybe-test: lua/error_handling_complex/test_pcall_varargs
-- origin: languages/lua/tests/lua/test_error_handling_complex.rs

local __w1 = "10"
local __i = 0

local function sum(a, b, c, d)
    return a + b + c + d
end
local ok, res = pcall(sum, 1, 2, 3, 4)
do local __t = tostring(res); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
