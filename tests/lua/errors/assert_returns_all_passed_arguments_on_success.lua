-- vybe-test: lua/errors/assert_returns_all_passed_arguments_on_success
-- origin: languages/lua/tests/lua/test_errors.rs

local __w1 = "10,err,30"
local __i = 0

local a, b, c = assert(10, "err", 30)
do local __t = tostring(a .. "," .. tostring(b) .. "," .. tostring(c)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
