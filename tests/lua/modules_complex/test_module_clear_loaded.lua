-- vybe-test: lua/modules_complex/test_module_clear_loaded
-- origin: languages/lua/tests/lua/test_modules_complex.rs

local __w1 = "val1 val2"
local __i = 0

package.loaded['temp_mod'] = 'val1'
        local r1 = require('temp_mod')
        package.loaded['temp_mod'] = nil
        local searcher = function(n) if n == 'temp_mod' then return function() return 'val2' end end end
        table.insert(package.searchers, 1, searcher)
        local r2 = require('temp_mod')
        table.remove(package.searchers, 1)
        do local __t = tostring(r1 .. ' ' .. r2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
