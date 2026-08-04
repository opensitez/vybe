-- vybe-test: lua/base_assert_default/test_assert_default_guarded
-- origin: languages/lua/tests/lua/test_base_assert_default.rs

local __w1 = "true\tnil"
local __i = 0

do local __t = tostring(select(1, pcall(function() assert(false or true) end))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
