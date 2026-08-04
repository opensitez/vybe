-- vybe-test: lua/iteration/next_called_manually_returns_pairs_in_sequence
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "a1"
local __i = 0

local t = {a = 1}
local k, v = next(t)
do local __t = tostring(k .. tostring(v)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
