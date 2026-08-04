-- vybe-test: lua/scoping/upvalue_written_before_closure_created_is_seen
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "10"
local __i = 0

local x
x = 10
local f = function() return x end
do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
