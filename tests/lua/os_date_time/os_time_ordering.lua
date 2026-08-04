-- vybe-test: lua/os_date_time/os_time_ordering
-- origin: languages/lua/tests/lua/test_os_date_time.rs

local __w1 = "true"
local __i = 0

local earlier = os.time({year=2000, month=1, day=1, hour=0, min=0, sec=0})
local later = os.time({year=2001, month=1, day=1, hour=0, min=0, sec=0})
do local __t = tostring(later > earlier); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
