lua_print! {
    test_package_path_basic => { "print(type(package.path))", "string" },
    test_package_cpath_basic => { "print(type(package.cpath))", "string" },
    test_package_loaded => { "print(type(package.loaded))", "table" },
    test_package_preload => { "print(type(package.preload))", "table" },
    test_package_searchers => { "print(type(package.searchers))", "table" },
    test_package_config => { "print(type(package.config))", "string" },
    test_package_searchpath_found => { "local path = package.searchpath('foo', '?.lua;?/init.lua'); print(type(path))", "string" },
    test_package_searchpath_not_found => { "local path, err = package.searchpath('foo', 'does_not_exist/?.lua'); print(tostring(path)..' '..tostring(type(err)=='string'))", "nil true" }
}
