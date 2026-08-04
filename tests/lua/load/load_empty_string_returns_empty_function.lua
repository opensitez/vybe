-- vybe-test: lua/load/load_empty_string_returns_empty_function
-- origin: languages/lua/tests/lua/test_load.rs

local __w1 = "true"
local __i = 0

local f = load("")
do local __t = tostring(type(f) == "function" and f() == nil); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
