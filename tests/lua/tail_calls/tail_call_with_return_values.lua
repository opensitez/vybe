-- vybe-test: lua/tail_calls/tail_call_with_return_values
-- origin: languages/lua/tests/lua/test_tail_calls.rs

local __w1 = "15"
local __i = 0

local function add(n, acc)
  if n == 0 then return acc end
  return add(n - 1, acc + n)
end
do local __t = tostring(add(5, 0)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
