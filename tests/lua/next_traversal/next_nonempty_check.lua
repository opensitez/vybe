-- vybe-test: lua/next_traversal/next_nonempty_check
-- origin: languages/lua/tests/lua/test_next_traversal.rs

local __w1 = "true"
local __i = 0

local function nonempty(t) return next(t) ~= nil end
do local __t = tostring(nonempty({1})); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
