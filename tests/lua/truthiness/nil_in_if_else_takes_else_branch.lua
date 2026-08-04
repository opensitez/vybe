-- vybe-test: lua/truthiness/nil_in_if_else_takes_else_branch
-- origin: languages/lua/tests/lua/test_truthiness.rs

local __w1 = "b"
local __i = 0

if nil then do local __t = tostring("a"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end else do local __t = tostring("b"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
