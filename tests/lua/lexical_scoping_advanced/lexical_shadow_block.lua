-- vybe-test: lua/lexical_scoping_advanced/lexical_shadow_block
-- origin: languages/lua/tests/lua/test_lexical_scoping_advanced.rs

local __w1 = "10\n5"
local __i = 0

local x = 5
if true then
  local x = 10
  do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
end
do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
