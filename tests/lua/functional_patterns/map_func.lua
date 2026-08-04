-- vybe-test: lua/functional_patterns/map_func
-- origin: languages/lua/tests/lua/test_functional_patterns.rs

local __w1 = "1,16"
local __i = 0

local function map(t, f)
  local r = {}
  for i, v in ipairs(t) do r[i] = f(v) end
  return r
end
local r = map({1,2,3,4}, function(x) return x * x end)
do local __t = tostring(r[1] .. "," .. r[4]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
