-- vybe-test: lua/metatables_concat/test_concat_fallback_error
-- origin: languages/lua/tests/lua/test_metatables_concat.rs

local __w1 = "false"
local __i = 0

local t1={}; local t2={}; local ok = pcall(function() return t1..t2 end); do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
