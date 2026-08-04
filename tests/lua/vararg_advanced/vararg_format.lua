-- vybe-test: lua/vararg_advanced/vararg_format
-- origin: languages/lua/tests/lua/test_vararg_advanced.rs

local __w1 = "2+3=5"
local __i = 0

local function fmt(pattern, ...)
  return string.format(pattern, ...)
end
do local __t = tostring(fmt("%d+%d=%d", 2, 3, 5)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
