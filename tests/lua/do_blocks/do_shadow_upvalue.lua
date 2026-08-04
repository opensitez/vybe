-- vybe-test: lua/do_blocks/do_shadow_upvalue
-- origin: languages/lua/tests/lua/test_do_blocks.rs

local __w1 = "3"
local __i = 0

local n = 1
do
  local n = 2
  do
    local n = 3
    do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
  end
end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
