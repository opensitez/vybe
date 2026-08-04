-- vybe-test: lua/functions_load/test_load_env
-- origin: languages/lua/tests/lua/test_functions_load.rs

local __w1 = "42"
local __i = 0

local env = {a=42}; local f = load('return a', 'chunk', 't', env); do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
