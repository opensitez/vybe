//! Dictionary lookup patterns: TryGetValue, GetValueOrDefault, ContainsKey guards, indexer contrast.

csharp_cases! {
    tryget_existing_string_key_returns_true_and_value => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["alpha"] = 10 }; Console.WriteLine(map.TryGetValue("alpha", out int v)); Console.WriteLine(v);"#,
        ["True", "10"]
    };

    tryget_missing_string_key_returns_false_and_zero => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int>(); Console.WriteLine(map.TryGetValue("ghost", out int v)); Console.WriteLine(v);"#,
        ["False", "0"]
    };

    tryget_out_var_infers_value_type => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["n"] = 7 }; if (map.TryGetValue("n", out var val)) Console.WriteLine(val);"#,
        ["7"]
    };

    tryget_false_branch_skips_out_value_use => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int>(); if (!map.TryGetValue("x", out var val)) Console.WriteLine("miss");"#,
        ["miss"]
    };

    tryget_after_indexer_insert_finds_new_entry => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int>(); map["fresh"] = 42; Console.WriteLine(map.TryGetValue("fresh", out int v)); Console.WriteLine(v);"#,
        ["True", "42"]
    };

    tryget_after_add_method_finds_added_pair => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int>(); map.Add("one", 1); Console.WriteLine(map.TryGetValue("one", out int v)); Console.WriteLine(v);"#,
        ["True", "1"]
    };

    tryget_after_remove_reports_absent_key => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["gone"] = 5 }; map.Remove("gone"); Console.WriteLine(map.TryGetValue("gone", out int v));"#,
        ["False"]
    };

    tryget_after_overwrite_reads_latest_value => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["k"] = 1 }; map["k"] = 9; map.TryGetValue("k", out int v); Console.WriteLine(v);"#,
        ["9"]
    };

    tryget_int_key_lookup_returns_stored_string => {
        r#"using System.Collections.Generic; var map = new Dictionary<int, string> { [42] = "answer" }; map.TryGetValue(42, out string s); Console.WriteLine(s);"#,
        ["answer"]
    };

    tryget_int_key_miss_returns_false => {
        r#"using System.Collections.Generic; var map = new Dictionary<int, string> { [1] = "a" }; Console.WriteLine(map.TryGetValue(99, out string s));"#,
        ["False"]
    };

    tryget_bool_key_stores_flag_value => {
        r#"using System.Collections.Generic; var map = new Dictionary<bool, string> { [true] = "yes" }; map.TryGetValue(true, out string s); Console.WriteLine(s);"#,
        ["yes"]
    };

    tryget_char_key_reads_single_character => {
        r#"using System.Collections.Generic; var map = new Dictionary<char, int> { ['A'] = 65 }; map.TryGetValue('A', out int code); Console.WriteLine(code);"#,
        ["65"]
    };

    get_value_or_default_missing_int_key_returns_zero => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int>(); Console.WriteLine(map.GetValueOrDefault("absent"));"#,
        ["0"]
    };

    get_value_or_default_existing_key_returns_stored_value => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["hit"] = 33 }; Console.WriteLine(map.GetValueOrDefault("hit"));"#,
        ["33"]
    };

    get_value_or_default_with_explicit_default_for_miss => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int>(); Console.WriteLine(map.GetValueOrDefault("nope", -1));"#,
        ["-1"]
    };

    get_value_or_default_with_explicit_default_ignores_on_hit => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["ok"] = 5 }; Console.WriteLine(map.GetValueOrDefault("ok", 99));"#,
        ["5"]
    };

    get_value_or_default_string_value_returns_null_for_miss => {
        r#"using System.Collections.Generic; var map = new Dictionary<int, string>(); Console.WriteLine(map.GetValueOrDefault(0) == null);"#,
        ["True"]
    };

    get_value_or_default_string_value_returns_payload_on_hit => {
        r#"using System.Collections.Generic; var map = new Dictionary<int, string> { [1] = "one" }; Console.WriteLine(map.GetValueOrDefault(1));"#,
        ["one"]
    };

    get_value_or_default_string_with_fallback_literal => {
        r#"using System.Collections.Generic; var map = new Dictionary<int, string>(); Console.WriteLine(map.GetValueOrDefault(2, "fallback"));"#,
        ["fallback"]
    };

    containskey_true_before_indexer_read => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["safe"] = 3 }; if (map.ContainsKey("safe")) Console.WriteLine(map["safe"]);"#,
        ["3"]
    };

    containskey_false_avoids_indexer_on_guard => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int>(); if (!map.ContainsKey("z")) Console.WriteLine("skip");"#,
        ["skip"]
    };

    containskey_and_tryget_agree_on_present_key => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["k"] = 8 }; Console.WriteLine(map.ContainsKey("k")); Console.WriteLine(map.TryGetValue("k", out int v));"#,
        ["True", "True"]
    };

    containskey_and_tryget_agree_on_absent_key => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["k"] = 8 }; Console.WriteLine(map.ContainsKey("z")); Console.WriteLine(map.TryGetValue("z", out int v));"#,
        ["False", "False"]
    };

    indexer_read_matches_tryget_for_same_key => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["x"] = 4 }; map.TryGetValue("x", out int viaTry); Console.WriteLine(map["x"] == viaTry);"#,
        ["True"]
    };

    tryget_preferred_over_indexer_for_optional_lookup => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1 }; int result = map.TryGetValue("b", out int v) ? v : -1; Console.WriteLine(result);"#,
        ["-1"]
    };

    tryget_ternary_selects_found_value => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 7 }; int result = map.TryGetValue("a", out int v) ? v : -1; Console.WriteLine(result);"#,
        ["7"]
    };

    get_value_or_default_matches_tryget_on_hit => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["p"] = 12 }; map.TryGetValue("p", out int t); Console.WriteLine(map.GetValueOrDefault("p") == t);"#,
        ["True"]
    };

    get_value_or_default_matches_default_on_miss => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int>(); map.TryGetValue("q", out int t); Console.WriteLine(map.GetValueOrDefault("q") == t);"#,
        ["True"]
    };

    tryget_twice_on_same_key_yields_same_value => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["dup"] = 6 }; map.TryGetValue("dup", out int a); map.TryGetValue("dup", out int b); Console.WriteLine(a == b);"#,
        ["True"]
    };

    tryget_after_clear_always_fails => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1 }; map.Clear(); Console.WriteLine(map.TryGetValue("a", out int v));"#,
        ["False"]
    };

    containskey_after_clear_is_false => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1 }; map.Clear(); Console.WriteLine(map.ContainsKey("a"));"#,
        ["False"]
    };

    tryget_readd_after_remove_finds_new_value => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["r"] = 1 }; map.Remove("r"); map["r"] = 2; map.TryGetValue("r", out int v); Console.WriteLine(v);"#,
        ["2"]
    };

    tryget_negative_int_key_lookup => {
        r#"using System.Collections.Generic; var map = new Dictionary<int, int> { [-1] = 100 }; map.TryGetValue(-1, out int v); Console.WriteLine(v);"#,
        ["100"]
    };

    tryget_zero_int_key_is_valid => {
        r#"using System.Collections.Generic; var map = new Dictionary<int, int> { [0] = 0 }; Console.WriteLine(map.TryGetValue(0, out int v)); Console.WriteLine(v);"#,
        ["True", "0"]
    };

    tryget_empty_string_key_is_valid => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { [""] = 1 }; map.TryGetValue("", out int v); Console.WriteLine(v);"#,
        ["1"]
    };

    tryget_long_string_key_roundtrip => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["longer_name"] = 55 }; map.TryGetValue("longer_name", out int v); Console.WriteLine(v);"#,
        ["55"]
    };

    tryget_distinct_keys_resolve_independently => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1, ["b"] = 2 }; map.TryGetValue("a", out int va); map.TryGetValue("b", out int vb); Console.WriteLine(va + vb);"#,
        ["3"]
    };

    get_value_or_default_after_overwrite => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["m"] = 1 }; map["m"] = 20; Console.WriteLine(map.GetValueOrDefault("m"));"#,
        ["20"]
    };

    get_value_or_default_negative_fallback => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["x"] = 5 }; Console.WriteLine(map.GetValueOrDefault("y", -99));"#,
        ["-99"]
    };

    containskey_distinguishes_similar_string_keys => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["ab"] = 1 }; Console.WriteLine(map.ContainsKey("ab")); Console.WriteLine(map.ContainsKey("a"));"#,
        ["True", "False"]
    };

    tryget_distinguishes_similar_string_keys => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["cat"] = 3 }; Console.WriteLine(map.TryGetValue("cat", out int v)); Console.WriteLine(map.TryGetValue("cats", out int w));"#,
        ["True", "False"]
    };

    tryget_out_string_default_is_null_on_miss => {
        r#"using System.Collections.Generic; var map = new Dictionary<int, string>(); map.TryGetValue(1, out string s); Console.WriteLine(s == null);"#,
        ["True"]
    };

    tryget_out_string_reads_on_hit => {
        r#"using System.Collections.Generic; var map = new Dictionary<int, string> { [3] = "three" }; map.TryGetValue(3, out string s); Console.WriteLine(s);"#,
        ["three"]
    };

    tryget_in_if_else_selects_found_branch => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["found"] = 11 }; if (map.TryGetValue("found", out int v)) Console.WriteLine("yes:" + v); else Console.WriteLine("no");"#,
        ["yes:11"]
    };

    tryget_in_if_else_selects_missing_branch => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int>(); if (map.TryGetValue("lost", out int v)) Console.WriteLine("yes"); else Console.WriteLine("no");"#,
        ["no"]
    };

    containskey_then_tryget_double_check_pattern => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["c"] = 9 }; int outVal = 0; if (map.ContainsKey("c") && map.TryGetValue("c", out outVal)) Console.WriteLine(outVal);"#,
        ["9"]
    };

    tryget_on_dictionary_with_three_entries => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1, ["b"] = 2, ["c"] = 3 }; map.TryGetValue("b", out int v); Console.WriteLine(v);"#,
        ["2"]
    };

    get_value_or_default_on_three_entry_map_miss => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1, ["b"] = 2, ["c"] = 3 }; Console.WriteLine(map.GetValueOrDefault("d", 0));"#,
        ["0"]
    };

    tryget_double_value_reads_fractional_number => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, double> { ["pi"] = 3.14 }; map.TryGetValue("pi", out double d); Console.WriteLine(d);"#,
        ["3.14"]
    };

    get_value_or_default_double_default_is_zero => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, double>(); Console.WriteLine(map.GetValueOrDefault("x"));"#,
        ["0"]
    };

    tryget_bool_stored_value => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, bool> { ["flag"] = true }; map.TryGetValue("flag", out bool b); Console.WriteLine(b);"#,
        ["True"]
    };

    containskey_after_overwrite_still_true => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["k"] = 1 }; map["k"] = 100; Console.WriteLine(map.ContainsKey("k"));"#,
        ["True"]
    };

    tryget_after_second_distinct_add => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int>(); map.Add("first", 1); map.Add("second", 2); map.TryGetValue("second", out int v); Console.WriteLine(v);"#,
        ["2"]
    };

    indexer_write_then_containskey_sees_new_key => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int>(); map["newkey"] = 77; Console.WriteLine(map.ContainsKey("newkey"));"#,
        ["True"]
    };

    tryget_and_get_value_or_default_both_succeed_on_hit => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["z"] = 44 }; bool ok = map.TryGetValue("z", out int t); int g = map.GetValueOrDefault("z"); Console.WriteLine(ok); Console.WriteLine(g);"#,
        ["True", "44"]
    };

    tryget_case_sensitive_string_key_distinction => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["Key"] = 1 }; Console.WriteLine(map.TryGetValue("Key", out int a)); Console.WriteLine(map.TryGetValue("key", out int b));"#,
        ["True", "False"]
    };

    containskey_case_sensitive_string_key_distinction => {
        r#"using System.Collections.Generic; var map = new Dictionary<string, int> { ["Mix"] = 2 }; Console.WriteLine(map.ContainsKey("Mix")); Console.WriteLine(map.ContainsKey("mix"));"#,
        ["True", "False"]
    };
}
