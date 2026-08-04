-- vybe-test: lua/os_library/os_date_with_format_year
-- origin: languages/lua/tests/lua/test_os_library.rs

local __w1 = "true"
local __i = 0

local s = os.date("%Y")
do local __t = tostring(#s == 4); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
