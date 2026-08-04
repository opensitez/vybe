-- vybe-test: lua/package_system/require_cache_mod
-- origin: languages/lua/tests/lua/test_package_system.rs

local __w1 = "99"
local __i = 0

package.preload["cached"] = function() return {n=0} end
local a = require("cached")
a.n = 99
local b = require("cached")
do local __t = tostring(b.n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
