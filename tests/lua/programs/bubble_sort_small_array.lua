-- vybe-test: lua/programs/bubble_sort_small_array
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "1,2,3"
local __i = 0

local t = {3, 1, 2}
for i = 1, #t - 1 do
  for j = 1, #t - i do
    if t[j] > t[j + 1] then t[j], t[j + 1] = t[j + 1], t[j] end
  end
end
do local __t = tostring(table.concat(t, ",")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
