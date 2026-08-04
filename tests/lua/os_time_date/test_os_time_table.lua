-- vybe-test: lua/os_time_date/test_os_time_table
-- origin: languages/lua/tests/lua/test_os_time_date.rs

local __w1 = "number"
local __i = 0

local t = os.time({year=2024, month=1, day=1, hour=12}); do local __t = tostring(type(t)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
