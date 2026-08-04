-- vybe-test: lua/io_library/io_write_to_stdout_returns_file_handle_or_nil
-- origin: languages/lua/tests/lua/test_io_library.rs

local r = io.write("")
print(r == io.stdout or r == nil or type(r) == "userdata")
