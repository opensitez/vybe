-- vybe-test: lua/os_time_date/test_os_rename
-- origin: languages/lua/tests/lua/test_os_time_date.rs

local __w1 = "nil true"
local __i = 0

local ok, err = os.rename('non_existent_A.tmp', 'non_existent_B.tmp'); do local __t = tostring(tostring(ok)..' '..tostring(type(err)=='string')); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
