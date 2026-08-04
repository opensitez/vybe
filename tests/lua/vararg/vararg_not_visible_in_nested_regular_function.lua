-- vybe-test: lua/vararg/vararg_not_visible_in_nested_regular_function
-- origin: languages/lua/tests/lua/test_vararg.rs

local __w1 = "0"
local __i = 0

local function outer(...)
  local function inner() return select('#', ...) end
  return inner()
end
do local __t = tostring(outer(1, 2, 3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
