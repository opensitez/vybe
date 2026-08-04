-- vybe-test: lua/io_implicit/test_io_write_implicit
-- origin: languages/lua/tests/lua/test_io_implicit.rs

local ok = pcall(function() io.write('') end); print(tostring(ok))
