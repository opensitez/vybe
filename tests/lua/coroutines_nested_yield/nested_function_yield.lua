-- vybe-test: lua/coroutines_nested_yield/nested_function_yield
-- origin: languages/lua/tests/lua/test_coroutines_nested_yield.rs

local __w1 = "inner,outer"
local __i = 0

local function inner() coroutine.yield("inner") end
local function outer() inner(); return "outer" end
local co = coroutine.create(outer)
local _, v1 = coroutine.resume(co)
local _, v2 = coroutine.resume(co)
do local __t = tostring(v1 .. "," .. v2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
