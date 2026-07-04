lua_print! {
    test_os_exit_exists => { "print(type(os.exit))", "function" },
    test_os_execute_exists => { "print(type(os.execute))", "function" },
    test_os_tmpname => { "local n = os.tmpname(); print(type(n))", "string" },
    test_os_tmpname_unique => { "local n1 = os.tmpname(); local n2 = os.tmpname(); print(tostring(n1 ~= n2))", "true" },
    test_os_setlocale_all => { "local l = os.setlocale('C', 'all'); print(type(l))", "string" }
}
