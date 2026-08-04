-- vybe-test: lua/vararg_advanced/vararg_forward
-- origin: languages/lua/tests/lua/test_vararg_advanced.rs

local __w1 = "7"
local __i = 0

local function add(a, b) return a + b end
local function proxy(...) return add(...) end
do local __t = tostring(proxy(3, 4)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
