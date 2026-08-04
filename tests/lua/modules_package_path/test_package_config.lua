-- vybe-test: lua/modules_package_path/test_package_config
-- origin: languages/lua/tests/lua/test_modules_package_path.rs

local __w1 = "string"
local __i = 0

do local __t = tostring(type(package.config)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
