-- vybe-test: lua/basics/nested_do_blocks_scoping_all_output
-- origin: languages/lua/tests/lua/test_basics.rs

local __w1 = "3"
local __i = 0

local a = 1
do
  do
    local a = 3
    do local __t = tostring(a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
  end
  do local __t = tostring(a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
