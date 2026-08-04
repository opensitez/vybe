-- vybe-test: lua/scoping/for_generic_control_vars_are_local_to_body
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "outer"
local __i = 0

local i = 'outer'
for i, v in ipairs({'a', 'b'}) do end
do local __t = tostring(i); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
