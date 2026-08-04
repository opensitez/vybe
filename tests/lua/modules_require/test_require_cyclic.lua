-- vybe-test: lua/modules_require/test_require_cyclic
-- origin: languages/lua/tests/lua/test_modules_require.rs

local __w1 = "init"
local __i = 0

package.loaded['cycle'] = 'init'; package.searchers[#package.searchers+1] = function(name) if name=='cycle' then return function() return require('cycle') end end end; do local __t = tostring(require('cycle')); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
