-- vybe-test: lua/functions_dump/test_dump_c_function_error
-- origin: languages/lua/tests/lua/test_functions_dump.rs

local __w1 = "false"
local __i = 0

local ok, err = pcall(function() string.dump(print) end); do local __t = tostring(tostring(ok)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
