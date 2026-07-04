lua_print! {
    test_gc_metamethod_called => { "local finalized = false; local t = setmetatable({}, {__gc = function() finalized = true end}); t = nil; collectgarbage('collect'); print(tostring(finalized))", "true" },
    test_gc_metamethod_error_ignored => { "local t = setmetatable({}, {__gc = function() error('boom') end}); t = nil; local ok = pcall(function() collectgarbage('collect') end); print(tostring(ok))", "true" },
    test_gc_metamethod_resurrection => { "local res; local t = setmetatable({a=42}, {__gc = function(obj) res = obj end}); t = nil; collectgarbage('collect'); print(res.a)", "42" },
    test_gc_string_no_finalizer => { "local ok = pcall(function() debug.setmetatable('', {__gc = function() end}) end); print(tostring(ok))", "true" }
}
