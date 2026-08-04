-- vybe-test: lua/metatables_len/test_len_error_no_metamethod_on_userdata
-- origin: languages/lua/tests/lua/test_metatables_len.rs

local __w1 = "false"
local __i = 0

local t=nil; local ok = pcall(function() return #t end); do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
