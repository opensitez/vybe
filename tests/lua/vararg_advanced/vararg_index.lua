-- vybe-test: lua/vararg_advanced/vararg_index
-- origin: languages/lua/tests/lua/test_vararg_advanced.rs

local __w1 = "20"
local __i = 0

local function f(...) local t = {...}; return t[2] end
do local __t = tostring(f(10, 20, 30)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
