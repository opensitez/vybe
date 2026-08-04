-- vybe-test: lua/package_system/require_preload_mod
-- origin: languages/lua/tests/lua/test_package_system.rs

local __w1 = "42"
local __i = 0

package.preload["mymod"] = function() return {val=42} end
local m = require("mymod")
do local __t = tostring(m.val); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
