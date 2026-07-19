//! Complex bitwise operations with shifts and negative values (Lua 5.3+ §3.4.2)

lua_print! {
    bitwise_not_neg_one => {
        "print(~(-1))\n",
        "0"
    },
    bitwise_shift_neg => {
        "local ok = pcall(function() return 1 << -1 end)\nprint(ok)\n",
        "true"
    },
    bitwise_and_mask => {
        "print(0xFFFF & 0x00FF)\n",
        "255"
    },
    bitwise_or_mask => {
        "print(0xF000 | 0x0F00)\n",
        "65280"
    },
    bitwise_xor_mask => {
        "print(0xAAAA ~ 0x5555)\n",
        "65535"
    },
}
