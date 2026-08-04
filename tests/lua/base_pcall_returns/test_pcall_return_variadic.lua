-- vybe-test: lua/base_pcall_returns/test_pcall_return_variadic
-- origin: languages/lua/tests/lua/test_base_pcall_returns.rs

local __w1 = "true"
local __i = 0

local function f(...)
  local a,b = ...
  return a * b
end
local ok, v = pcall(f, 3, 4)
do local __t = tostring(ok and v == 12); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
