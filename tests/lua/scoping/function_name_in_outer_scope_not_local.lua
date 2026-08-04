-- vybe-test: lua/scoping/function_name_in_outer_scope_not_local
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "1"
local __i = 0

function h() return 1 end
do local __t = tostring(h()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
