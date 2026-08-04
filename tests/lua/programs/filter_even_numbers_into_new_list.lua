-- vybe-test: lua/programs/filter_even_numbers_into_new_list
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "2"
local __i = 0

local src = {1, 2, 3, 4, 5}
local out = {}
for _, v in ipairs(src) do if v % 2 == 0 then table.insert(out, v) end end
do local __t = tostring(#out); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
