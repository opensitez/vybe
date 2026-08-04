-- vybe-test: lua/os_time_date/test_os_getenv
-- origin: languages/lua/tests/lua/test_os_time_date.rs

local __w1 = "true"
local __i = 0

local e = os.getenv('PATH'); do local __t = tostring(type(e) == 'string' or e == nil); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
