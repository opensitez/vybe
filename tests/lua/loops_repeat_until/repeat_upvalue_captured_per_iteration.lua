-- vybe-test: lua/loops_repeat_until/repeat_upvalue_captured_per_iteration
-- origin: languages/lua/tests/lua/test_loops_repeat_until.rs

local __w1 = "10,20,30,"
local __i = 0

local out = ''
local i = 0
repeat
  i = i + 1
  local x = i * 10
  out = out .. x .. ','
until i == 3
do local __t = tostring(out); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
