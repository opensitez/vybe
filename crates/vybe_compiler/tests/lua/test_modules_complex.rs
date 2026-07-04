lua_print! {
    test_module_return_table => {
        "local m = {}
m.val = 42
function m.get() return m.val end
package.loaded['my_module'] = m
local req = require('my_module')
print(req.get())",
        "42"
    },
    test_module_custom_searcher => {
        "local searcher = function(modname)
            if modname == 'custom_mod' then
                return function() return 'custom result' end
            end
        end
        table.insert(package.searchers, 1, searcher)
        local res = require('custom_mod')
        table.remove(package.searchers, 1)
        print(res)",
        "custom result"
    },
    test_module_cyclic_dependency => {
        "local m1, m2 = {}, {}
        package.loaded['m1'] = m1
        package.loaded['m2'] = m2
        m1.get_m2 = function() return require('m2') end
        m2.get_m1 = function() return require('m1') end
        print(tostring(m1.get_m2() == m2) .. ' ' .. tostring(m2.get_m1() == m1))",
        "true true"
    },
    test_module_require_error => {
        "local searcher = function(modname)
            if modname == 'error_mod' then
                return function() error('module load error') end
            end
        end
        table.insert(package.searchers, 1, searcher)
        local ok, err = pcall(require, 'error_mod')
        table.remove(package.searchers, 1)
        print(tostring(ok) .. ' ' .. tostring(string.find(err, 'module load error') ~= nil))",
        "false true"
    },
    test_module_clear_loaded => {
        "package.loaded['temp_mod'] = 'val1'
        local r1 = require('temp_mod')
        package.loaded['temp_mod'] = nil
        local searcher = function(n) if n == 'temp_mod' then return function() return 'val2' end end end
        table.insert(package.searchers, 1, searcher)
        local r2 = require('temp_mod')
        table.remove(package.searchers, 1)
        print(r1 .. ' ' .. r2)",
        "val1 val2"
    }
}
