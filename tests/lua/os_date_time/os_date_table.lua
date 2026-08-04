-- vybe-test: lua/os_date_time/os_date_table
-- origin: languages/lua/tests/lua/test_os_date_time.rs

local __w1 = "2024,3,10"
local __i = 0

local epoch = os.time({year=2024, month=3, day=10, hour=0, min=0, sec=0})
local d = os.date("*t", epoch)
do local __t = tostring(d.year .. "," .. d.month .. "," .. d.day); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
