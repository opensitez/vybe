-- vybe-test: lua/error_handling_pcall/test_pcall_error_object
-- origin: languages/lua/tests/lua/test_error_handling_pcall.rs

local __w1 = "true"
local __i = 0

local t={}; local ok, err = pcall(function() error(t) end); do local __t = tostring(tostring(err==t)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
