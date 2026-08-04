-- vybe-test: lua/programs/dutch_flag_partition_zeros_and_ones
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "0,0,1,1,1"
local __i = 0

local t = {1, 0, 1, 0, 1}
local low, high = 1, #t
local i = 1
while i <= high do
  if t[i] == 0 then
    t[low], t[i] = t[i], t[low]
    low = low + 1
    i = i + 1
  else
    high = high - 1
    t[i], t[high] = t[high], t[i]
  end
end
do local __t = tostring(table.concat(t, ",")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
