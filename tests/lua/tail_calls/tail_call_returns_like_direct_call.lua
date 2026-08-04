-- vybe-test: lua/tail_calls/tail_call_returns_like_direct_call
-- origin: languages/lua/tests/lua/test_tail_calls.rs

local __w1 = "ok"
local __i = 0

local function tail(n)
  if n == 0 then return "ok" end
  return tail(n - 1)
end
do local __t = tostring(tail(3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
