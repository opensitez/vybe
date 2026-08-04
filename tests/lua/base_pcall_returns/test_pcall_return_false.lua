-- vybe-test: lua/base_pcall_returns/test_pcall_return_false
-- origin: languages/lua/tests/lua/test_base_pcall_returns.rs

local __w1 = "true"
local __i = 0

local ok, v = pcall(function() return false end)
do local __t = tostring(ok and (v == false)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
