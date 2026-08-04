-- vybe-test: lua/package_system/sys_package_preload_tbl
-- origin: languages/lua/tests/lua/test_package_system.rs

local __w1 = "table"
local __i = 0

do local __t = tostring(type(package.preload)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
