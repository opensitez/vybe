-- vybe-test: lua/vararg/vararg_in_middle_of_expression_list_adjusted_to_one
-- origin: languages/lua/tests/lua/test_vararg.rs

local __w1 = "1\t99\tnil"
local __i = 0

local function f(...) return ..., 99 end
local a, b, c = f(1, 2, 3)
do local __t = tostring(a) .. "\t" .. tostring(b) .. "\t" .. tostring(c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
