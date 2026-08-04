-- vybe-test: lua/programs/zip_two_arrays_into_pairs_table
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "2y"
local __i = 0

local a, b = {1, 2}, {"x", "y"}
local z = {}
for i = 1, math.min(#a, #b) do z[i] = a[i] .. b[i] end
do local __t = tostring(z[2]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
