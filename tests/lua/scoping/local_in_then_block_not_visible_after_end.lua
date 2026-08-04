-- vybe-test: lua/scoping/local_in_then_block_not_visible_after_end
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "nil"
local __i = 0

if true then
  local secret = 'yes'
end
do local __t = tostring(tostring(secret)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
