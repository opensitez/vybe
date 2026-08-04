-- vybe-test: lua/iteration/break_in_pairs_loop_early
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "2"
local __i = 0

local t = {a=1, b=2, c=3}
local n = 0
for _ in pairs(t) do
  n = n + 1
  if n == 2 then break end
end
do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
