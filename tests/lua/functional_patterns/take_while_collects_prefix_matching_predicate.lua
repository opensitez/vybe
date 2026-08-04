-- vybe-test: lua/functional_patterns/take_while_collects_prefix_matching_predicate
-- origin: languages/lua/tests/lua/test_functional_patterns.rs

local __w1 = "1,2,3"
local __i = 0

local function take_while(t, pred)
  local r = {}
  for _, v in ipairs(t) do
    if not pred(v) then break end
    r[#r+1] = v
  end
  return r
end
local t = take_while({1,2,3,4,5}, function(x) return x < 4 end)
do local __t = tostring(table.concat(t, ',')); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
