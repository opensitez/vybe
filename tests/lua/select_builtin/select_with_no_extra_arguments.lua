-- vybe-test: lua/select_builtin/select_with_no_extra_arguments
-- origin: languages/lua/tests/lua/test_select_builtin.rs

local __w1 = "nil,nil"
local __i = 0

local a, b = select(1)
do local __t = tostring(tostring(a) .. "," .. tostring(b)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
