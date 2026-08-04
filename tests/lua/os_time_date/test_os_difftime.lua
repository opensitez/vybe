-- vybe-test: lua/os_time_date/test_os_difftime
-- origin: languages/lua/tests/lua/test_os_time_date.rs

local __w1 = "50.0"
local __i = 0

local diff = os.difftime(100, 50); do local __t = tostring(diff); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
