-- vybe-test: lua/functions/tail_call_does_not_grow_stack_is_semantics
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "done"
local __i = 0

local function tail(n)
  if n == 0 then return "done" end
  return tail(n - 1)
end
do local __t = tostring(tail(5)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
