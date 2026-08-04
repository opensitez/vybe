-- vybe-test: lua/package/loaded_table_reuses_same_table
-- origin: languages/lua/tests/lua/test_package.rs

local __w1 = "2"
local __i = 0

package.loaded.mod = {x = 1}
local a = require("mod")
a.x = 2
local b = require("mod")
do local __t = tostring(b.x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
