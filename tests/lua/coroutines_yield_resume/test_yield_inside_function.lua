-- vybe-test: lua/coroutines_yield_resume/test_yield_inside_function
-- origin: languages/lua/tests/lua/test_coroutines_yield_resume.rs

local __w1 = "99 100"
local __i = 0

local function f() coroutine.yield(99) end; local co = coroutine.create(function() f() return 100 end); local ok, v1 = coroutine.resume(co); local ok2, v2 = coroutine.resume(co); do local __t = tostring(v1..' '..v2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
