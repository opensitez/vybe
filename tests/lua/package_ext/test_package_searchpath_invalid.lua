-- vybe-test: lua/package_ext/test_package_searchpath_invalid
-- origin: languages/lua/tests/lua/test_package_ext.rs

local __w1 = "nil true"
local __i = 0

local p, err = package.searchpath('foo', './does_not_exist/?.lua'); do local __t = tostring(tostring(p)..' '..tostring(type(err)=='string')); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
