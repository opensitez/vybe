lua_print! {
    test_rep_basic => { "print(string.rep('a', 3))", "aaa" },
    test_rep_zero => { "print(string.rep('a', 0))", "" },
    test_rep_negative => { "print(string.rep('a', -1))", "" },
    test_rep_with_sep => { "print(string.rep('a', 3, ','))", "a,a,a" },
    test_rep_with_sep_one => { "print(string.rep('a', 1, ','))", "a" },
    test_rep_with_sep_zero => { "print(string.rep('a', 0, ','))", "" },
    test_rep_empty_string => { "print(string.rep('', 10, ','))", ",,,,,,,,," },
    test_rev_basic => { "print(string.reverse('abc'))", "cba" },
    test_rev_empty => { "print(string.reverse(''))", "" },
    test_rev_palindrome => { "print(string.reverse('racecar'))", "racecar" },
    test_lower_basic => { "print(string.lower('AbCd!'))", "abcd!" },
    test_upper_basic => { "print(string.upper('AbCd!'))", "ABCD!" }
}
