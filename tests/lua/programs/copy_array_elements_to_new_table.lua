-- vybe-test: lua/programs/copy_array_elements_to_new_table
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "2"
local __i = 0

local src = {1, 2}
local dst = {}
for i = 1, #src do dst[i] = src[i] end
do local __t = tostring(dst[2]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
