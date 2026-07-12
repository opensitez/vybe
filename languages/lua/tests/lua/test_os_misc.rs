lua_print! {
    test_os_exit_exists => { "print(type(os.exit))", "function" },
    test_os_execute_exists => { "print(type(os.execute))", "function" },
    test_os_tmpname => { "local n = os.tmpname(); print(type(n))", "string" },
    test_os_tmpname_unique => { "local n1 = os.tmpname(); local n2 = os.tmpname(); print(tostring(n1 ~= n2))", "true" },
    test_os_setlocale_all => { "local l = os.setlocale('C', 'all'); print(type(l))", "string" },
    os_getenv_returns_string_or_nil => {
        "local val = os.getenv(\"PATH\")\nprint(val == nil or type(val) == \"string\")\n",
        "true"
    },
    os_getenv_nonexistent_returns_nil => {
        "local val = os.getenv(\"SOME_NONEXISTENT_VAR_XYZ_123\")\nprint(tostring(val))\n",
        "nil"
    },
    os_clock_returns_number => {
        "local c = os.clock()\nprint(type(c))\n",
        "number"
    },
    os_difftime_computes_difference => {
        "local d = os.difftime(100, 50)\nprint(d)\n",
        "50"
    },
    os_remove_nonexistent_fails_and_returns_nil_plus_error => {
        "local ok, err = os.remove(\"nonexistent_file_xyz_123.txt\")\nprint(ok == nil and type(err) == \"string\")\n",
        "true"
    },
    os_rename_nonexistent_fails_and_returns_nil_plus_error => {
        "local ok, err = os.rename(\"nonexistent1.txt\", \"nonexistent2.txt\")\nprint(ok == nil and type(err) == \"string\")\n",
        "true"
    },
    os_setlocale_invalid_returns_nil => {
        "local res = os.setlocale(\"invalid_locale_name_xyz\")\nprint(tostring(res))\n",
        "nil"
    },
    os_execute_with_command_returns_termination_status => {
        "local ok, status, code = os.execute(\"true\")\nprint(type(ok) == \"boolean\" or ok == nil)\n",
        "true"
    },
}
