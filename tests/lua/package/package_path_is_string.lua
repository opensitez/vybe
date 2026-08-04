-- vybe-test: lua/package/package_path_is_string
-- origin: languages/lua/tests/lua/test_package.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(type(package.path) == "string"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
