-- vybe-test: lua/coroutines/coroutine_create_requires_function_argument
-- origin: languages/lua/tests/lua/test_coroutines.rs

local __w1 = "false"
local __i = 0

local ok = pcall(coroutine.create, 1)
do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
