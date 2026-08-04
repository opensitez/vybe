-- vybe-test: lua/functions_multiple_returns/test_multret_trailing_in_function_call
-- origin: languages/lua/tests/lua/test_functions_multiple_returns.rs

local __w1 = "123"
local __i = 0

local function f() return 2, 3 end; local function g(a,b,c) return a..b..c end; do local __t = tostring(g(1, f())); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
