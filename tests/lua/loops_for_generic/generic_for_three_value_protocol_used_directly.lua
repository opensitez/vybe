-- vybe-test: lua/loops_for_generic/generic_for_three_value_protocol_used_directly
-- origin: languages/lua/tests/lua/test_loops_for_generic.rs

local __w1 = "15"
local __i = 0

local function stateless(state, i)
  i = i + 1
  if i <= state then return i end
end
local sum = 0
for i in stateless, 5, 0 do sum = sum + i end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
