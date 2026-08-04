-- vybe-test: lua/do_blocks/do_block_redefines_global_as_local
-- origin: languages/lua/tests/lua/test_do_blocks.rs

local __w1 = "20\n10"
local __i = 0

global_var_xyz = 10
do
  local global_var_xyz = 20
  do local __t = tostring(global_var_xyz); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
end
do local __t = tostring(global_var_xyz); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
