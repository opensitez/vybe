-- vybe-test: lua/functional_patterns/drop_while_removes_prefix_matching_predicate
-- origin: languages/lua/tests/lua/test_functional_patterns.rs

local __w1 = "3,4,5"
local __i = 0

local function drop_while(t, pred)
  local i = 1
  while i <= #t and pred(t[i]) do i = i + 1 end
  local r = {}
  for j = i, #t do r[#r+1] = t[j] end
  return r
end
local t = drop_while({1,2,3,4,5}, function(x) return x < 3 end)
do local __t = tostring(table.concat(t, ',')); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
