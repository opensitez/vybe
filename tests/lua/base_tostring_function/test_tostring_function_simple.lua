-- vybe-test: lua/base_tostring_function/test_tostring_function_simple
-- origin: languages/lua/tests/lua/test_base_tostring_function.rs

local __w1 = "true"
local __i = 0

local f = function() return 1 + 1 end; local s = tostring(f); do local __t = tostring(type(s) == "string" and string.sub(s, 1, 9) == "function:"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
