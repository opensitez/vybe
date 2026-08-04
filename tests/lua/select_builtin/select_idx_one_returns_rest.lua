-- vybe-test: lua/select_builtin/select_idx_one_returns_rest
-- origin: languages/lua/tests/lua/test_select_builtin.rs

local __w1 = "a\tb\tc"
local __i = 0

do local __t = tostring(select(1, 'a', 'b', 'c')); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
