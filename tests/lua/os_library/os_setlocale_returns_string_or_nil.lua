-- vybe-test: lua/os_library/os_setlocale_returns_string_or_nil
-- origin: languages/lua/tests/lua/test_os_library.rs

local __w1 = "true"
local __i = 0

local r = os.setlocale("C")
do local __t = tostring(r == "C" or r == nil); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
