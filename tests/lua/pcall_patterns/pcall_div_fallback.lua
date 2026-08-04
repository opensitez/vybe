-- vybe-test: lua/pcall_patterns/pcall_div_fallback
-- origin: languages/lua/tests/lua/test_pcall_patterns.rs

local __w1 = "-1"
local __i = 0

local function try_div(a, b)
  if b == 0 then error("div0") end
  return a / b
end
local ok, v = pcall(try_div, 10, 0)
local result = ok and v or -1
do local __t = tostring(result); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
