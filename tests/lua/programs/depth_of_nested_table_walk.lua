-- vybe-test: lua/programs/depth_of_nested_table_walk
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "2"
local __i = 0

local t = {a = {b = {c = 1}}}
local depth = 0
local node = t
while node.a do depth = depth + 1 node = node.a end
do local __t = tostring(depth); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
