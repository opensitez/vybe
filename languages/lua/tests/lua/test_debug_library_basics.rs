//! debug library: getinfo, traceback, getlocal (Lua 5.x §6.10)

lua_print! {
    getinfo_tbl => {
        "local info = debug.getinfo(1)\nprint(type(info))\n",
        "table"
    },
    getinfo_source => {
        "local info = debug.getinfo(1, \"S\")\nprint(type(info.source))\n",
        "string"
    },
    getinfo_what => {
        "local function f() return debug.getinfo(1, \"S\").what end\nprint(f())\n",
        "Lua"
    },
    getinfo_line => {
        "local info = debug.getinfo(1, \"l\")\nprint(info.currentline > 0)\n",
        "true"
    },
    getinfo_name => {
        "local function myFunc() return debug.getinfo(1, \"n\").name end\nprint(myFunc())\n",
        "myFunc"
    },
    traceback_str => {
        "local tb = debug.traceback()\nprint(type(tb))\n",
        "string"
    },
    traceback_msg => {
        "local tb = debug.traceback(\"error here\")\nprint(type(tb))\n",
        "string"
    },
    getlocal_val => {
        "local function f()\n  local x = 42\n  local name, val = debug.getlocal(1, 1)\n  print(name .. \"=\" .. val)\nend\nf()\n",
        "x=42"
    },
    getlocal_nil => {
        "local name, val = debug.getlocal(1, 100)\nprint(tostring(name))\n",
        "nil"
    },
}
