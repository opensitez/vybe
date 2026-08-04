-- vybe-test: lua/coroutines_advanced/test_coroutine_yielding_across_pcall
-- origin: languages/lua/tests/lua/test_coroutines_advanced.rs

local __w1 = "42 false true"
local __i = 0

local co = coroutine.create(function()
    local ok, err = pcall(function()
        coroutine.yield(42)
        error('boom')
    end)
    return ok, err
end)
local ok, res = coroutine.resume(co)
local ok2, ok_inner, err = coroutine.resume(co)
do local __t = tostring(res .. ' ' .. tostring(ok_inner) .. ' ' .. tostring(string.find(err, 'boom') ~= nil)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
