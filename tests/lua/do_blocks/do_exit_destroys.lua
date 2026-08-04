-- vybe-test: lua/do_blocks/do_exit_destroys
-- origin: languages/lua/tests/lua/test_do_blocks.rs

local __w1 = "true"
local __i = 0

do
  local tmp = 100
end
local exists = (tmp == nil)
do local __t = tostring(exists); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
