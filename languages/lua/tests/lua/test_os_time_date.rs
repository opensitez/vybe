lua_print! {
    test_os_time_no_args => { "local t = os.time(); print(type(t))", "number" },
    test_os_time_table => { "local t = os.time({year=2024, month=1, day=1, hour=12}); print(type(t))", "number" },
    test_os_date_string => { "local d = os.date('%Y-%m-%d', 1704067200); print(type(d))", "string" },
    test_os_date_table => { "local d = os.date('*t', 1704067200); print(d.year)", "2024" },
    test_os_difftime => { "local diff = os.difftime(100, 50); print(diff)", "50.0" },
    test_os_clock => { "local c = os.clock(); print(type(c))", "number" },
    test_os_getenv => { "local e = os.getenv('PATH'); print(type(e) == 'string' or e == nil)", "true" },
    test_os_setlocale => { "local l = os.setlocale('C'); print(type(l) == 'string' or l == nil)", "true" },
    test_os_remove => { "local ok, err = os.remove('non_existent_file_12345.tmp'); print(tostring(ok)..' '..tostring(type(err)=='string'))", "nil true" },
    test_os_rename => { "local ok, err = os.rename('non_existent_A.tmp', 'non_existent_B.tmp'); print(tostring(ok)..' '..tostring(type(err)=='string'))", "nil true" }
}
