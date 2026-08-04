-- vybe-test: lua/vararg/vararg_with_extra_args_past_select_index
-- origin: languages/lua/tests/lua/test_vararg.rs

local __w1 = "20,30"
local __i = 0

local function from_second(...)
  return select(2, ...)
end
local a, b = from_second(10, 20, 30)
do local __t = tostring(a .. ',' .. b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
