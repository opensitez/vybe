-- vybe-test: lua/package/require_caches_module_in_loaded
-- origin: languages/lua/tests/lua/test_package.rs

local __w1 = "1"
local __i = 0

package.loaded.fake = {v = 1}
local m = require("fake")
do local __t = tostring(m.v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
