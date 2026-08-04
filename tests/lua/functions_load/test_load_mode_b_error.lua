-- vybe-test: lua/functions_load/test_load_mode_b_error
-- origin: languages/lua/tests/lua/test_functions_load.rs

local __w1 = "nil true"
local __i = 0

local f, err = load('return 1', 'chunk', 'b'); do local __t = tostring(tostring(f)..' '..tostring(type(err)=='string')); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
