-- vybe-test: lua/load_custom_env/load_fn_def
-- origin: languages/lua/tests/lua/test_load_custom_env.rs

local __w1 = "25"
local __i = 0

local f = load("return function(n) return n * n end")
do local __t = tostring(f()(5)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
