-- vybe-test: lua/functions/function_receives_table_and_mutates_it
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "7,7,7"
local __i = 0

local function fill(t, val)
  for i = 1, 3 do t[i] = val end
end
local data = {}
fill(data, 7)
do local __t = tostring(data[1] .. ',' .. data[2] .. ',' .. data[3]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
