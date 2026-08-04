-- vybe-test: lua/error_handling_pcall/test_pcall_non_function
-- origin: languages/lua/tests/lua/test_error_handling_pcall.rs

local __w1 = "false"
local __i = 0

local ok, err = pcall(42); do local __t = tostring(tostring(ok)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
