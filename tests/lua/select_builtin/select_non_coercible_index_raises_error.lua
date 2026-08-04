-- vybe-test: lua/select_builtin/select_non_coercible_index_raises_error
-- origin: languages/lua/tests/lua/test_select_builtin.rs

local __w1 = "false"
local __i = 0

local ok = pcall(select, {}, 1, 2)
do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
