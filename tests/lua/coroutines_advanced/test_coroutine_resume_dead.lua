-- vybe-test: lua/coroutines_advanced/test_coroutine_resume_dead
-- origin: languages/lua/tests/lua/test_coroutines_advanced.rs

local __w1 = "false true"
local __i = 0

local co = coroutine.create(function() return 1 end)
coroutine.resume(co)
local ok, err = coroutine.resume(co)
do local __t = tostring(tostring(ok) .. ' ' .. tostring(string.find(err, 'dead coroutine') ~= nil)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
