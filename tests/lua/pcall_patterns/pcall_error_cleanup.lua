-- vybe-test: lua/pcall_patterns/pcall_error_cleanup
-- origin: languages/lua/tests/lua/test_pcall_patterns.rs

local __w1 = "true"
local __i = 0

local cleaned = false
local ok = pcall(function()
  error("boom")
end)
cleaned = true
do local __t = tostring(cleaned); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
