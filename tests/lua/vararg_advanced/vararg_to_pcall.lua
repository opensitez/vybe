-- vybe-test: lua/vararg_advanced/vararg_to_pcall
-- origin: languages/lua/tests/lua/test_vararg_advanced.rs

local __w1 = "true\t15"
local __i = 0

local function f(a, b) return a + b end
local ok, v = pcall(f, 7, 8)
do local __t = tostring(ok) .. "\t" .. tostring(v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
