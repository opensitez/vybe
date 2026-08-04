-- vybe-test: lua/vararg_functions_advanced/vararg_select_count
-- origin: languages/lua/tests/lua/test_vararg_functions_advanced.rs

local __w1 = "3"
local __i = 0

local function f(...) return select("#", ...) end
do local __t = tostring(f(nil, nil, 3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
