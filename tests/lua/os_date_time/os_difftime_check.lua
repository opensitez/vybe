-- vybe-test: lua/os_date_time/os_difftime_check
-- origin: languages/lua/tests/lua/test_os_date_time.rs

local __w1 = "5.0"
local __i = 0

local t1 = os.time({year=2000, month=1, day=1, hour=0, min=0, sec=0})
local t2 = os.time({year=2000, month=1, day=1, hour=0, min=0, sec=5})
do local __t = tostring(os.difftime(t2, t1)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
