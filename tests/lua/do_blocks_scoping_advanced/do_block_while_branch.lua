-- vybe-test: lua/do_blocks_scoping_advanced/do_block_while_branch
-- origin: languages/lua/tests/lua/test_do_blocks_scoping_advanced.rs

local __w1 = "6"
local __i = 0

local x = 1
while x < 2 do
  local y = 5
  do
    local z = x + y
    do local __t = tostring(z); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
  end
  x = x + 1
end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
