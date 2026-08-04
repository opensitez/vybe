-- vybe-test: lua/coercion/tostring_for_building_messages
-- origin: languages/lua/tests/lua/test_coercion.rs

local __w1 = "count=7"
local __i = 0

local n = 7
do local __t = tostring("count=" .. tostring(n)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
