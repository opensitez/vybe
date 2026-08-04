-- vybe-test: lua/pcall_patterns/pcall_retry_loop
-- origin: languages/lua/tests/lua/test_pcall_patterns.rs

local __w1 = "3"
local __i = 0

local attempts = 0
local ok = false
while not ok and attempts < 3 do
  attempts = attempts + 1
  ok = pcall(function()
    if attempts < 3 then error("retry") end
  end)
end
do local __t = tostring(attempts); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
