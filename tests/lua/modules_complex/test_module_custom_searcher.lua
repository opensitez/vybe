-- vybe-test: lua/modules_complex/test_module_custom_searcher
-- origin: languages/lua/tests/lua/test_modules_complex.rs

local __w1 = "custom result"
local __i = 0

local searcher = function(modname)
            if modname == 'custom_mod' then
                return function() return 'custom result' end
            end
        end
        table.insert(package.searchers, 1, searcher)
        local res = require('custom_mod')
        table.remove(package.searchers, 1)
        do local __t = tostring(res); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
