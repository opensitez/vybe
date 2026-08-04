-- vybe-test: lua/package_path_searching/searchpath_missing
-- origin: languages/lua/tests/lua/test_package_path_searching.rs

local __w1 = "nil"
local __i = 0

local p = package.searchpath and package.searchpath("missing_module", package.path)
do local __t = tostring(tostring(p)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
