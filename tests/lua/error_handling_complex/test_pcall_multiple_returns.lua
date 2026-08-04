-- vybe-test: lua/error_handling_complex/test_pcall_multiple_returns
-- origin: languages/lua/tests/lua/test_error_handling_complex.rs

local __w1 = "true 1 2 3"
local __i = 0

local function multi() return 1, 2, 3 end
local a,b,c,d = pcall(multi)
do local __t = tostring(tostring(a) .. ' ' .. b .. ' ' .. c .. ' ' .. d); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
