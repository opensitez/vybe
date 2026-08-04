-- vybe-test: lua/base_pcall_truths/test_pcall_variadic_wrapper
-- origin: languages/lua/tests/lua/test_base_pcall_truths.rs

local __w1 = "true"
local __i = 0

local function f(...)
  local a,b,c = ...
  return a+b+c
end
local ok, v = pcall(f, 1,2,3)
do local __t = tostring(ok and v == 6); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
