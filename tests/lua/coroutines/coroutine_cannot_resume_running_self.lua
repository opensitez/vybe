-- vybe-test: lua/coroutines/coroutine_cannot_resume_running_self
-- origin: languages/lua/tests/lua/test_coroutines.rs

local __w1 = "false,cannot resume running coroutine"
local __i = 0

local co
co = coroutine.create(function()
  local ok, err = coroutine.resume(co)
  do local __t = tostring(ok .. "," .. tostring(err)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
end)
coroutine.resume(co)

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
