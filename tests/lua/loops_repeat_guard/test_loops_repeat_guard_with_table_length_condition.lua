-- vybe-test: lua/loops_repeat_guard/test_loops_repeat_guard_with_table_length_condition
-- origin: languages/lua/tests/lua/test_loops_repeat_guard.rs

local __w1 = "1"
local __i = 0

local t = {}
repeat table.insert(t, 1) until #t == 1
do local __t = tostring(#t); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
