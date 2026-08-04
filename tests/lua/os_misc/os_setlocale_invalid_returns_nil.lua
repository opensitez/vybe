-- vybe-test: lua/os_misc/os_setlocale_invalid_returns_nil
-- origin: languages/lua/tests/lua/test_os_misc.rs

local __w1 = "nil"
local __i = 0

local res = os.setlocale("invalid_locale_name_xyz")
do local __t = tostring(tostring(res)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
