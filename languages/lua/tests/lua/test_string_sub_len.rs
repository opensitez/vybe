lua_print! {
    test_sub_basic => { "print(string.sub('hello', 2, 4))", "ell" },
    test_sub_default_end => { "print(string.sub('hello', 2))", "ello" },
    test_sub_negative_start => { "print(string.sub('hello', -3))", "llo" },
    test_sub_negative_end => { "print(string.sub('hello', 1, -2))", "hell" },
    test_sub_out_of_bounds => { "print(string.sub('hello', 10, 20))", "" },
    test_sub_reversed_range => { "print(string.sub('hello', 4, 2))", "" },
    test_len_basic => { "print(string.len('hello'))", "5" },
    test_len_empty => { "print(string.len(''))", "0" },
    test_len_embedded_zeros => { "print(string.len('a\\0b\\0c'))", "5" }
}
