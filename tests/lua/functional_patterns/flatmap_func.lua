-- vybe-test: lua/functional_patterns/flatmap_func
-- origin: languages/lua/tests/lua/test_functional_patterns.rs

local __w1 = "1,10,2,20,3,30"
local __i = 0

local function flatmap(t, f)
  local r = {}
  for _, v in ipairs(t) do
    for _, w in ipairs(f(v)) do r[#r+1] = w end
  end
  return r
end
local r = flatmap({1,2,3}, function(x) return {x, x*10} end)
do local __t = tostring(table.concat(r, ",")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
