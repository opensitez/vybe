-- vybe-test: lua/coroutine_wrappers/test_wrap_yield_resume
-- origin: languages/lua/tests/lua/test_coroutine_wrappers.rs

local __w1 = "1 2 3"
local __i = 0

local f = coroutine.wrap(function()
    coroutine.yield(1)
    coroutine.yield(2)
    return 3
end)
do local __t = tostring(f() .. ' ' .. f() .. ' ' .. f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
