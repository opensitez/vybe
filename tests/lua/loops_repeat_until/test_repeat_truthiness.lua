-- vybe-test: lua/loops_repeat_until/test_repeat_truthiness
-- origin: languages/lua/tests/lua/test_loops_repeat_until.rs

local __w1 = "12"
local __i = 0

local s=''; local t={1, 2, nil}; local i=1; repeat s=s..t[i]; i=i+1 until not t[i]; do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
