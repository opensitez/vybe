-- vybe-test: lua/scoping/global_written_inside_function_visible_outside
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "9"
local __i = 0

function setk() key = 9 end
setk()
do local __t = tostring(key); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
