lua_print! {
    test_package_searchpath_valid => { "local p = package.searchpath('foo', './?.lua'); print(p)", "./foo.lua" },
    test_package_searchpath_invalid => { "local p, err = package.searchpath('foo', './does_not_exist/?.lua'); print(tostring(p)..' '..tostring(type(err)=='string'))", "nil true" },
    test_package_searchpath_replace_sep => { "local p = package.searchpath('foo.bar', './?.lua', '.', '/'); print(p)", "./foo/bar.lua" },
    test_package_searchpath_replace_rep => { "local p = package.searchpath('foo', './?.lua', '.', '/', 'x'); print(p)", "./foo.lua" },
    test_package_loadlib_not_found => { "local f, err = package.loadlib('does_not_exist.so', 'foo'); print(tostring(f)..' '..tostring(type(err)=='string'))", "nil true" },
    test_package_config_exists => { "print(type(package.config))", "string" }
}
