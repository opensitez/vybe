-- vybe-test: lua/scoping_blocks/test_block_local_function
-- origin: languages/lua/tests/lua/test_scoping_blocks.rs

local __w1 = "42"
local __i = 0

local f1; do local function f2() return 42 end; f1 = f2 end; do local __t = tostring(f1()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
