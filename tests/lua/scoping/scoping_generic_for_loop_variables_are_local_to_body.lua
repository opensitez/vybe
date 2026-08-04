-- vybe-test: lua/scoping/scoping_generic_for_loop_variables_are_local_to_body
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "outer"
local __i = 0

local k = "outer"
local t = {key = "val"}
for k, v in pairs(t) do end
do local __t = tostring(k); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
