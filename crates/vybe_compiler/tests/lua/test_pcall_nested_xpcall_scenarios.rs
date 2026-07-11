//! Advanced pcall/xpcall combination scenarios and custom error models (Lua 5.x §6.1)

lua_print! {
    xpcall_in_pcall_ok => {
        "local ok, inner_ok, val = pcall(function()\n  return xpcall(function() return \"ok\" end, function() end)\nend)\nprint(ok, inner_ok, val)\n",
        "true\ttrue\tok"
    },
    xpcall_in_pcall_err => {
        "local ok, inner_ok, val = pcall(function()\n  return xpcall(function() error(\"fail\", 0) end, function(e) return \"handled:\"..e end)\nend)\nprint(ok, inner_ok, val)\n",
        "true\tfalse\thandled:fail"
    },
    pcall_in_xpcall => {
        "local function handler(e) return \"h:\" .. e end\nlocal ok, val = xpcall(function()\n  local ok2, res = pcall(function() error(\"inner\") end)\n  return ok2, res\nend, handler)\nprint(ok, val)\n",
        "true\tfalse\tinput:5: inner"
    },
    xpcall_custom_class => {
        "local function handler(e) return {msg = e.msg, code = e.code} end\nlocal ok, err = xpcall(function() error({msg=\"bad\", code=500}) end, handler)\nprint(ok, err.msg, err.code)\n",
        "false\tbad\t500"
    },
}
