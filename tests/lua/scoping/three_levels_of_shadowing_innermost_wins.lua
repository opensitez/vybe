-- vybe-test: lua/scoping/three_levels_of_shadowing_innermost_wins
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "inner\nmiddle\nouter"
local __i = 0

local x = 'outer'
do
  local x = 'middle'
  do
    local x = 'inner'
    do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
  end
  do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
end
do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
