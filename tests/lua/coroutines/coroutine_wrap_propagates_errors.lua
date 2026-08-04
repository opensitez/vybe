-- vybe-test: lua/coroutines/coroutine_wrap_propagates_errors
-- origin: languages/lua/tests/lua/test_coroutines.rs

local __w1 = "false,wrap_fail"
local __i = 0

local f = coroutine.wrap(function() error("wrap_fail") end)
local ok, err = pcall(f)
do local __t = tostring(ok .. "," .. tostring(err:match("wrap_fail"))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
