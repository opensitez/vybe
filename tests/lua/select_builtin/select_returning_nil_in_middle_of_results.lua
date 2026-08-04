-- vybe-test: lua/select_builtin/select_returning_nil_in_middle_of_results
-- origin: languages/lua/tests/lua/test_select_builtin.rs

local __w1 = "nil,30,nil"
local __i = 0

local a, b, c = select(2, 10, nil, 30)
do local __t = tostring(tostring(a) .. "," .. tostring(b) .. "," .. tostring(c)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
