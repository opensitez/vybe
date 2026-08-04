-- vybe-test: lua/functions_load/test_load_mode_t
-- origin: languages/lua/tests/lua/test_functions_load.rs

local __w1 = "1"
local __i = 0

local f = load('return 1', 'chunk', 't'); do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
