-- vybe-test: lua/io_file_handles/test_io_write_stdout
-- origin: languages/lua/tests/lua/test_io_file_handles.rs

local ok, err = io.stdout:write(''); print(tostring(ok))
