//! `goto` and labels — Lua 5.2+ manual §3.3.8.

lua_print! {
    goto_skips_unreachable_assignment => {
        "local x = 1\n goto finish\n x = 2\n ::finish::\n print(x)\n",
        "1"
    },
    goto_can_form_simple_loop => {
        "local n = 0\n::loop::\n n = n + 1\n if n < 3 then goto loop end\n print(n)\n",
        "3"
    },
    goto_forward_to_label => {
        "local s = \"\"\n goto two\n s = s .. \"1\"\n ::two::\n s = s .. \"2\"\n print(s)\n",
        "2"
    },
    break_and_goto_do_not_share_labels => {
        "local n = 0\nwhile n < 2 do\n  n = n + 1\n  goto done\nend\nn = 9\n::done::\nprint(n)\n",
        "1"
    },
}
