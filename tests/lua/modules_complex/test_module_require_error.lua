-- vybe-test: lua/modules_complex/test_module_require_error
-- origin: languages/lua/tests/lua/test_modules_complex.rs

local __w1 = "false true"
local __i = 0

local searcher = function(modname)
            if modname == 'error_mod' then
                return function() error('module load error') end
            end
        end
        table.insert(package.searchers, 1, searcher)
        local ok, err = pcall(require, 'error_mod')
        table.remove(package.searchers, 1)
        do local __t = tostring(tostring(ok) .. ' ' .. tostring(string.find(err, 'module load error') ~= nil)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
