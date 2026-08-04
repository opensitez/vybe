-- vybe-test: lua/do_blocks/do_nested_shares_upvalue
-- origin: languages/lua/tests/lua/test_do_blocks.rs

local __w1 = "2"
local __i = 0

local count = 0
do
  count = count + 1
  do
    count = count + 1
  end
end
do local __t = tostring(count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
