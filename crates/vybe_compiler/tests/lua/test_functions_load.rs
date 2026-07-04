lua_print! {
    test_load_string_valid => { "local f = load('return 42'); print(f())", "42" },
    test_load_string_invalid => { "local f, err = load('return %'); print(tostring(f)..' '..tostring(type(err)=='string'))", "nil true" },
    test_load_function_valid => { "local s='return 42'; local i=1; local f = load(function() local chunk = string.sub(s, i, i); i=i+1; if chunk == '' then return nil else return chunk end end); print(f())", "42" },
    test_load_chunkname => { "local f, err = load('error()', 'mychunk'); local ok, err_msg = pcall(f); print(tostring(string.find(err_msg, 'mychunk') ~= nil))", "true" },
    test_load_mode_t => { "local f = load('return 1', 'chunk', 't'); print(f())", "1" },
    test_load_mode_b_error => { "local f, err = load('return 1', 'chunk', 'b'); print(tostring(f)..' '..tostring(type(err)=='string'))", "nil true" },
    test_load_env => { "local env = {a=42}; local f = load('return a', 'chunk', 't', env); print(f())", "42" },
    test_loadfile_not_found => { "local f, err = loadfile('does_not_exist_file.lua'); print(tostring(f)..' '..tostring(type(err)=='string'))", "nil true" }
}
