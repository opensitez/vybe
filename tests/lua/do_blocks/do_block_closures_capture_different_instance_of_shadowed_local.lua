-- vybe-test: lua/do_blocks/do_block_closures_capture_different_instance_of_shadowed_local
-- origin: languages/lua/tests/lua/test_do_blocks.rs

local __w1 = "10,20"
local __i = 0

local f1, f2
local x = 10
f1 = function() return x end
do
  local x = 20
  f2 = function() return x end
end
do local __t = tostring(f1() .. "," .. f2()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
