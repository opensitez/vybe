-- vybe-test: lua/load/loadstring_alias_when_present
-- origin: languages/lua/tests/lua/test_load.rs

local __w1 = "3"
local __i = 0

local loader = loadstring or load
local f = loader("return 3")
do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
