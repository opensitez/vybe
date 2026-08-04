-- vybe-test: lua/base_pcall_truths/test_pcall_multiple_return_second
-- origin: languages/lua/tests/lua/test_base_pcall_truths.rs

local __w1 = "1"
local __i = 0

local ok, a = pcall(function() return 1, 2 end)
do local __t = tostring(ok and a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
