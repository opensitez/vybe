-- vybe-test: lua/load_custom_env/load_syntax_error
-- origin: languages/lua/tests/lua/test_load_custom_env.rs

local __w1 = "true\tstring"
local __i = 0

local f, err = load("return @")
do local __t = tostring(f == nil) .. "\t" .. tostring(type(err)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
