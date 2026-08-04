-- vybe-test: lua/package_path_searching/searchers_count
-- origin: languages/lua/tests/lua/test_package_path_searching.rs

local __w1 = "true"
local __i = 0

local searchers = package.searchers or package.loaders
do local __t = tostring(#searchers >= 1); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
