-- vybe-test: lua/next_traversal/next_traversal_loop
-- origin: languages/lua/tests/lua/test_next_traversal.rs

local __w1 = "2"
local __i = 0

local t = {p=1, q=2}
local seen = 0
local k = nil
repeat
  k, _ = next(t, k)
  if k then seen = seen + 1 end
until not k
do local __t = tostring(seen); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
