-- vybe-test: lua/functions_tailcalls/test_tailcall_mutually_recursive
-- origin: languages/lua/tests/lua/test_functions_tailcalls.rs

local __w1 = "g f"
local __i = 0

local f, g; f = function(n) if n==0 then return 'f' else return g(n-1) end end; g = function(n) if n==0 then return 'g' else return f(n-1) end end; do local __t = tostring(f(9)..' '..f(10)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
