-- vybe-test: lua/functional_patterns/group_by_partitions_into_table_of_lists
-- origin: languages/lua/tests/lua/test_functional_patterns.rs

local __w1 = "3,3"
local __i = 0

local function group_by(t, key_fn)
  local groups = {}
  for _, v in ipairs(t) do
    local k = key_fn(v)
    if not groups[k] then groups[k] = {} end
    groups[k][#groups[k]+1] = v
  end
  return groups
end
local g = group_by({1,2,3,4,5,6}, function(x) return x % 2 == 0 and 'even' or 'odd' end)
do local __t = tostring(#g.even .. ',' .. #g.odd); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
