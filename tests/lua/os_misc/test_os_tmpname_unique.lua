-- vybe-test: lua/os_misc/test_os_tmpname_unique
-- origin: languages/lua/tests/lua/test_os_misc.rs

local __w1 = "true"
local __i = 0

local n1 = os.tmpname(); local n2 = os.tmpname(); do local __t = tostring(tostring(n1 ~= n2)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
