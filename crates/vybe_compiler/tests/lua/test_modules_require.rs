lua_print! {
    test_require_basic => { "package.loaded['mymod'] = 42; print(require('mymod'))", "42" },
    test_require_caching => { "local c=0; package.searchers[#package.searchers+1] = function(name) if name=='testmod' then return function() c=c+1; return c end end end; require('testmod'); require('testmod'); print(c)", "1" },
    test_require_return_true_if_no_return => { "package.searchers[#package.searchers+1] = function(name) if name=='testmod2' then return function() end end end; print(require('testmod2'))", "true" },
    test_require_not_found_error => { "local ok, err = pcall(function() require('nonexistent_module') end); print(tostring(ok)..' '..tostring(string.find(err, 'module') ~= nil))", "false true" },
    test_require_cyclic => { "package.loaded['cycle'] = 'init'; package.searchers[#package.searchers+1] = function(name) if name=='cycle' then return function() return require('cycle') end end end; print(require('cycle'))", "init" }
}
