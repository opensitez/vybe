-- vybe-test: lua/do_blocks_scoping_advanced/do_block_for_branch
-- origin: languages/lua/tests/lua/test_do_blocks_scoping_advanced.rs

local __w1 = "30"
local __i = 0

local s = 0
for i = 1, 2 do
  do
    local x = i * 10
    s = s + x
  end
end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
