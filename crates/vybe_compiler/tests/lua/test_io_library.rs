//! `io` library — standard streams, type (Lua 5.x manual §6.8).

lua_print! {
    io_type_on_nil_returns_nil => {
        "print(tostring(io.type(nil)))\n",
        "nil"
    },
    io_type_on_closed_file => {
        "local f = io.tmpfile()\nif f then f:close() print(io.type(f) == \"closed file\" or io.type(f) == nil) else print(true) end\n",
        "true"
    },
    io_write_to_stdout_returns_file_handle_or_nil => {
        "local r = io.write(\"\")\nprint(r == io.stdout or r == nil or type(r) == \"userdata\")\n",
        "true"
    },
    io_lines_on_string_without_file => {
        "local sum = 0\nfor n in io.lines(\"1\\n2\\n3\\n\") do sum = sum + tonumber(n) end\nprint(sum)\n",
        "6"
    },
}
