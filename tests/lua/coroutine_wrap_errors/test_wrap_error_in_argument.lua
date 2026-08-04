-- vybe-test: lua/coroutine_wrap_errors/test_wrap_error_in_argument
-- origin: languages/lua/tests/lua/test_coroutine_wrap_errors.rs

local __w1 = "true"
local __i = 0

local function bad(x) if x == 0 then error("bad") end return x end
local f = coroutine.wrap(bad)
local ok, err = pcall(function() f(0) end)
do local __t = tostring(ok == false and string.find(err, "bad") ~= nil); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
