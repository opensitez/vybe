//! Goto statements jumping across nested block boundaries and loop control (Lua 5.x §3.3.4)

lua_print! {
    goto_nested_do_exit => {
        "local reached = false\ndo\n  do\n    goto target\n  end\n  reached = true\nend\n::target::\nprint(reached)\n",
        "false"
    },
    goto_while_loop_break => {
        "local n = 0\nwhile n < 5 do\n  n = n + 1\n  if n == 3 then goto exit_loop end\nend\n::exit_loop::\nprint(n)\n",
        "3"
    },
    goto_for_loop_break => {
        "local last = 0\nfor i = 1, 10 do\n  if i == 4 then goto done end\n  last = i\nend\n::done::\nprint(last)\n",
        "3"
    },
    goto_skip_local_decl => {
        "local ok = true\ngoto target\nlocal val = 42\n::target::\nprint(ok)\n",
        "true"
    },
}
