-- vybe-test: lua/os_date_time/os_date_month
-- origin: languages/lua/tests/lua/test_os_date_time.rs

local __w1 = "06"
local __i = 0

local epoch = os.time({year=2024, month=6, day=15, hour=12, min=0, sec=0})
do local __t = tostring(os.date("%m", epoch)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
