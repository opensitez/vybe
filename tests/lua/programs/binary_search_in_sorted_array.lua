-- vybe-test: lua/programs/binary_search_in_sorted_array
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "4"
local __i = 0

local t = {1, 3, 5, 7, 9}
local target, lo, hi = 7, 1, #t
while lo <= hi do
  local mid = math.floor((lo + hi) / 2)
  if t[mid] == target then do local __t = tostring(mid); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end return end
  if t[mid] < target then lo = mid + 1 else hi = mid - 1 end
end
do local __t = tostring(0); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
