-- vybe-test: lua/functions_multiple_returns/test_multret_middle_in_return
-- origin: languages/lua/tests/lua/test_functions_multiple_returns.rs

local __w1 = "13nil"
local __i = 0

local function f() return 1, 2 end; local function g() return f(), 3 end; local a,b,c = g(); do local __t = tostring(a..b..tostring(c)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
