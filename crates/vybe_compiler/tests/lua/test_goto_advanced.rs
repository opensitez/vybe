//! goto statements — forward jumps, loops, scoping limits (Lua 5.x §3.3.4)

lua_print! {
    goto_skips => {
        "goto skip\nprint(\"skipped\")\n::skip::\nprint(\"reached\")\n",
        "reached"
    },
    goto_if_end => {
        "local x = 5\nif x > 3 then goto done end\nprint(\"not reached\")\n::done::\nprint(x)\n",
        "5"
    },
    goto_continue_while => {
        "local i = 0\nlocal s = 0\n::again::\ni = i + 1\nif i > 5 then goto done end\nif i % 2 == 0 then goto again end\ns = s + i\ngoto again\n::done::\nprint(s)\n",
        "9"
    },
    goto_continue_for => {
        "local s = 0\nfor i = 1, 6 do\n  if i % 2 == 0 then goto continue end\n  s = s + i\n  ::continue::\nend\nprint(s)\n",
        "9"
    },
    goto_nested_do => {
        "do\n  do\n    goto out\n  end\n  print(\"inner\")\nend\n::out::\nprint(\"out\")\n",
        "out"
    },
    goto_decl_label => {
        "local x = 1\ngoto lbl\nx = 2\n::lbl::\nprint(x)\n",
        "1"
    },
    goto_multi_labels => {
        "local step = 1\ngoto step2\n::step1::\nstep = 10\ngoto done\n::step2::\nstep = 2\ngoto done\n::done::\nprint(step)\n",
        "2"
    },
}
