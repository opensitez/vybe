-- vybe-test: lua/modules_complex/test_module_return_table
-- origin: languages/lua/tests/lua/test_modules_complex.rs

local __w1 = "42"
local __i = 0

local m = {}
m.val = 42
function m.get() return m.val end
package.loaded['my_module'] = m
local req = require('my_module')
do local __t = tostring(req.get()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
