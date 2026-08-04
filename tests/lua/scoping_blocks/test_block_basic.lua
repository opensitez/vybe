-- vybe-test: lua/scoping_blocks/test_block_basic
-- origin: languages/lua/tests/lua/test_scoping_blocks.rs

local __w1 = "1"
local __i = 0

local a=1; do local a=2 end; do local __t = tostring(a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
