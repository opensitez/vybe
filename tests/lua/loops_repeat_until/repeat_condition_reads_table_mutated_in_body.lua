-- vybe-test: lua/loops_repeat_until/repeat_condition_reads_table_mutated_in_body
-- origin: languages/lua/tests/lua/test_loops_repeat_until.rs

local __w1 = "4"
local __i = 0

local t = {done = false}
local n = 0
repeat
  n = n + 1
  if n == 4 then t.done = true end
until t.done
do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
