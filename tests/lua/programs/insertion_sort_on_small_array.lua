-- vybe-test: lua/programs/insertion_sort_on_small_array
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "1,2,3,4"
local __i = 0

local t = {3, 1, 4, 2}
for i = 2, #t do
  local key = t[i]
  local j = i - 1
  while j > 0 and t[j] > key do
    t[j + 1] = t[j]
    j = j - 1
  end
  t[j + 1] = key
end
do local __t = tostring(table.concat(t, ",")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
