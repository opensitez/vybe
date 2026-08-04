-- vybe-test: lua/io_file_handles/test_io_flush_stdout
-- origin: languages/lua/tests/lua/test_io_file_handles.rs

local ok = io.stdout:flush(); print(tostring(ok))
