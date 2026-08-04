-- vybe-test: lua/functional_patterns/zip_func
-- origin: languages/lua/tests/lua/test_functional_patterns.rs

local __w1 = "2b"
local __i = 0

local function zip(a, b)
  local r = {}
  for i = 1, math.min(#a, #b) do r[i] = {a[i], b[i]} end
  return r
end
local z = zip({1,2,3}, {"a","b","c"})
do local __t = tostring(z[2][1] .. z[2][2]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
