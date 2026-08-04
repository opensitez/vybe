-- vybe-test: lua/scoping/local_not_initialized_starts_as_nil
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "nil"
local __i = 0

local uninit
do local __t = tostring(tostring(uninit)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
