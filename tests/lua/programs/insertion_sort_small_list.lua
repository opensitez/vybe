-- vybe-test: lua/programs/insertion_sort_small_list
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "1"
local __i = 0

local t = {5, 2, 4, 1}
for i = 2, #t do
  local key, j = t[i], i - 1
  while j >= 1 and t[j] > key do t[j + 1] = t[j] j = j - 1 end
  t[j + 1] = key
end
do local __t = tostring(t[1]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
