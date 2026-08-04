-- vybe-test: lua/modules_require/test_require_basic
-- origin: languages/lua/tests/lua/test_modules_require.rs

local __w1 = "42"
local __i = 0

package.loaded['mymod'] = 42; do local __t = tostring(require('mymod')); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
