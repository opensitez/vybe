-- vybe-test: lua/coroutine_status_initial/test_status_create_multiple_threads
-- origin: languages/lua/tests/lua/test_coroutine_status_initial.rs

local __w1 = "suspended/suspended"
local __i = 0

local a = coroutine.create(function() return 1 end)
local b = coroutine.create(function() return 2 end)
do local __t = tostring(coroutine.status(a) .. "/" .. coroutine.status(b)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
