-- vybe-test: lua/os_date_time/os_time_type
-- origin: languages/lua/tests/lua/test_os_date_time.rs

local __w1 = "number"
local __i = 0

do local __t = tostring(type(os.time())); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
