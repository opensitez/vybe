-- vybe-test: lua/functional_patterns/filter_func
-- origin: languages/lua/tests/lua/test_functional_patterns.rs

local __w1 = "3,2"
local __i = 0

local function filter(t, pred)
  local r = {}
  for _, v in ipairs(t) do
    if pred(v) then r[#r+1] = v end
  end
  return r
end
local evens = filter({1,2,3,4,5,6}, function(x) return x % 2 == 0 end)
do local __t = tostring(#evens .. "," .. evens[1]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
