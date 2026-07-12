//! Advanced do...end blocks inside nested logical branches (Lua 5.x §3.3.2)

lua_print! {
    do_block_if_branch => {
        "local x = 5\nif true then\n  local y = 10\n  do\n    local z = x + y\n    print(z)\n  end\nend\n",
        "15"
    },
    do_block_while_branch => {
        "local x = 1\nwhile x < 2 do\n  local y = 5\n  do\n    local z = x + y\n    print(z)\n  end\n  x = x + 1\nend\n",
        "6"
    },
    do_block_for_branch => {
        "local s = 0\nfor i = 1, 2 do\n  do\n    local x = i * 10\n    s = s + x\n  end\nend\nprint(s)\n",
        "30"
    },
}
