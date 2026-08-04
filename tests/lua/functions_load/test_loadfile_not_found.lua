-- vybe-test: lua/functions_load/test_loadfile_not_found
-- origin: languages/lua/tests/lua/test_functions_load.rs

local __w1 = "nil true"
local __i = 0

local f, err = loadfile('does_not_exist_file.lua'); do local __t = tostring(tostring(f)..' '..tostring(type(err)=='string')); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
