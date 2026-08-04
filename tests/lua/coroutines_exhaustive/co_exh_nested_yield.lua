-- vybe-test: lua/coroutines_exhaustive/co_exh_nested_yield
-- origin: languages/lua/tests/lua/test_coroutines_exhaustive.rs

local __w1 = "inner\touter"
local __i = 0

local function inner() coroutine.yield("inner") end
local co = coroutine.create(function() inner(); return "outer" end)
local _, v1 = coroutine.resume(co)
local _, v2 = coroutine.resume(co)
do local __t = tostring(v1) .. "\t" .. tostring(v2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
