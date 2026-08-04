-- vybe-test: lua/vararg_advanced/vararg_recursive
-- origin: languages/lua/tests/lua/test_vararg_advanced.rs

local __w1 = "123"
local __i = 0

local function concat(...)
  local out = ''
  for i = 1, select('#', ...) do out = out .. tostring(select(i, ...)) end
  return out
end
do local __t = tostring(concat(1, 2, 3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
