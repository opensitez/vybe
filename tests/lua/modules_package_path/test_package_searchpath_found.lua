-- vybe-test: lua/modules_package_path/test_package_searchpath_found
-- origin: languages/lua/tests/lua/test_modules_package_path.rs

local __w1 = "string"
local __i = 0

local path = package.searchpath('foo', '?.lua;?/init.lua'); do local __t = tostring(type(path)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
