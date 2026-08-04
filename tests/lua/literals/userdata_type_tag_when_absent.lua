-- vybe-test: lua/literals/userdata_type_tag_when_absent
-- origin: languages/lua/tests/lua/test_literals.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(type(io.stdin) == "userdata" or type(io.stdin) == "nil"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
