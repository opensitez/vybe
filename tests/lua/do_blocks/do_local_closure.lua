-- vybe-test: lua/do_blocks/do_local_closure
-- origin: languages/lua/tests/lua/test_do_blocks.rs

local __w1 = "42"
local __i = 0

local fn
do
  local secret = 42
  fn = function() return secret end
end
do local __t = tostring(fn()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
