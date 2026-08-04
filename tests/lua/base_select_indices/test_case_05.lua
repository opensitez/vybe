-- vybe-test: lua/base_select_indices/test_case_05
-- origin: languages/lua/tests/lua/test_base_select_indices.rs

local __w1 = "true"
local __i = 0

local idx = 1; local ok, _ = pcall(function() return select(idx, 10, 20, 30) end); local expect = (idx == 1 or idx == 2 or idx == 3 or idx == -1 or idx == -2 or idx == -3); do local __t = tostring(ok == expect); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
