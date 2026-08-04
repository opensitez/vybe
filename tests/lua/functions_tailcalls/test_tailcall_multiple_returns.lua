-- vybe-test: lua/functions_tailcalls/test_tailcall_multiple_returns
-- origin: languages/lua/tests/lua/test_functions_tailcalls.rs

local __w1 = "1 2"
local __i = 0

local function f() return 1, 2 end; local function g() return f() end; local a, b = g(); do local __t = tostring(a..' '..b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
