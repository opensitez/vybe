-- vybe-test: lua/programs/short_circuit_and_skips_rhs
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "false"
local __i = 0

local called = false
local function side() called = true return false end
if false and side() then end
do local __t = tostring(tostring(called)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
