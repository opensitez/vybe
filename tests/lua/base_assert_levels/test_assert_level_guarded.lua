-- vybe-test: lua/base_assert_levels/test_assert_level_guarded
-- origin: languages/lua/tests/lua/test_base_assert_levels.rs

local __w1 = "true"
local __i = 0

local v1, v2, v3 = assert(14, 22, undefined); do local __t = tostring(v1 + v2 == 14 + 22); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
