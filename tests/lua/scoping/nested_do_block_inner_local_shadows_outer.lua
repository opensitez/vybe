-- vybe-test: lua/scoping/nested_do_block_inner_local_shadows_outer
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "2"
local __i = 0

local v = 1
 do
  local v = 2
  do local __t = tostring(v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
 end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
