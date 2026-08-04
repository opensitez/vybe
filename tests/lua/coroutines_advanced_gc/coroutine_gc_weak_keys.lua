-- vybe-test: lua/coroutines_advanced_gc/coroutine_gc_weak_keys
-- origin: languages/lua/tests/lua/test_coroutines_advanced_gc.rs

local __w1 = "42"
local __i = 0

local t = setmetatable({}, {__mode="k"})
local co = coroutine.create(function() end)
t[co] = 42
do local __t = tostring(t[co]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
