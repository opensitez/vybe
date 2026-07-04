lua_print! {
    test_io_input_get => { "print(io.type(io.input()))", "file" },
    test_io_output_get => { "print(io.type(io.output()))", "file" },
    test_io_write_implicit => { "local ok = pcall(function() io.write('') end); print(tostring(ok))", "true" },
    test_io_flush_implicit => { "local ok = pcall(function() io.flush() end); print(tostring(ok))", "true" },
    test_io_read_implicit_type => { "local ok, err = pcall(function() io.read(true) end); print(tostring(ok))", "false" },
    test_io_input_set_invalid => { "local ok, err = pcall(function() io.input(true) end); print(tostring(ok))", "false" },
    test_io_output_set_invalid => { "local ok, err = pcall(function() io.output(true) end); print(tostring(ok))", "false" }
}
