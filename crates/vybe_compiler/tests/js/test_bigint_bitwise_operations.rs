//! BigInt arithmetic, bitwise ops, comparison, and conversion edge cases.

crate::js_cases! {
    bigint_addition => {
        r#"console.log((1n+2n).toString());"#,
        ["3"]
    };

    bigint_subtraction => {
        r#"console.log((5n-3n).toString());"#,
        ["2"]
    };

    bigint_multiplication => {
        r#"console.log((6n*7n).toString());"#,
        ["42"]
    };

    bigint_division_truncates => {
        r#"console.log((10n/3n).toString());"#,
        ["3"]
    };

    bigint_remainder => {
        r#"console.log((10n%3n).toString());"#,
        ["1"]
    };

    bigint_exponentiation => {
        r#"console.log((2n**10n).toString());"#,
        ["1024"]
    };

    bigint_unary_minus => {
        r#"console.log((-5n).toString());"#,
        ["-5"]
    };

    bigint_bitwise_and => {
        r#"console.log((12n&10n).toString());"#,
        ["8"]
    };

    bigint_bitwise_or => {
        r#"console.log((12n|10n).toString());"#,
        ["14"]
    };

    bigint_bitwise_xor => {
        r#"console.log((12n^10n).toString());"#,
        ["6"]
    };

    bigint_left_shift => {
        r#"console.log((1n<<4n).toString());"#,
        ["16"]
    };

    bigint_right_shift => {
        r#"console.log((16n>>2n).toString());"#,
        ["4"]
    };

    bigint_less_than => {
        r#"console.log(1n<2n);"#,
        ["true"]
    };

    bigint_equality => {
        r#"console.log(7n===7n);"#,
        ["true"]
    };

    bigint_not_equal => {
        r#"console.log(1n!==2n);"#,
        ["true"]
    };

    bigint_greater_than_or_equal => {
        r#"console.log(5n>=5n);"#,
        ["true"]
    };

    bigint_typeof => {
        r#"console.log(typeof 1n);"#,
        ["bigint"]
    };

    bigint_constructor_from_decimal_string => {
        r#"console.log(BigInt("99"));"#,
        ["99n"]
    };

    bigint_constructor_from_hex_string => {
        r#"console.log(BigInt("0xff"));"#,
        ["255n"]
    };

    bigint_constructor_from_binary_string => {
        r#"console.log(BigInt("0b1010"));"#,
        ["10n"]
    };

    bigint_to_string_decimal => {
        r#"console.log((255n).toString());"#,
        ["255"]
    };

    bigint_to_string_hex => {
        r#"console.log((255n).toString(16));"#,
        ["ff"]
    };

    bigint_mixed_with_number_add_throws => {
        r#"try{console.log(1n+1);}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    // Node-verified: relational comparison across BigInt/Number is LEGAL
    // (§7.2.13) — only arithmetic mixing throws. 1n < 1 is false.
    bigint_mixed_with_number_compare_throws => {
        r#"try{console.log(1n<1);}catch(e){console.log(e instanceof TypeError);}"#,
        ["false"]
    };

    bigint_zero_division_throws => {
        r#"try{console.log(1n/0n);}catch(e){console.log(e instanceof RangeError);}"#,
        ["true"]
    };

    bigint_as_object_property => {
        r#"const o={v:9n}; console.log(o.v);"#,
        ["9n"]
    };

    bigint_in_array_map => {
        r#"console.log([1n,2n].map(x=>x+1n).join(","));"#,
        ["2,3"]
    };

    bigint_json_stringify_throws => {
        r#"try{JSON.stringify({v:1n});}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

    bigint_max_safe_increment => {
        r#"const b=9007199254740991n; console.log((b+1n).toString());"#,
        ["9007199254740992"]
    };

    bigint_negative_value => {
        r#"console.log((-42n).toString());"#,
        ["-42"]
    };

    // Node-verified: `b++` on a BigInt binding is perfectly legal (yields
    // 2n) — nothing here is a SyntaxError.
    bigint_postfix_increment_syntax_error => {
        r#"let b=1n; b++; console.log(b);"#,
        ["2n"]
    };

    bigint_valueof_returns_same => {
        r#"const b=3n; console.log(b.valueOf()===b);"#,
        ["true"]
    };

    bigint_in_switch_strict_match => {
        r#"const v=2n; let r=""; switch(v){case 2n:r="ok";break;default:r="no";} console.log(r);"#,
        ["ok"]
    };

    bigint_modulo_with_negative => {
        r#"console.log((-10n%3n).toString());"#,
        ["-1"]
    };

    bigint_bitwise_not => {
        r#"console.log((~0n).toString());"#,
        ["-1"]
    };

    bigint_division_negative => {
        r#"console.log((-7n/2n).toString());"#,
        ["-3"]
    };

    bigint_chain_arithmetic => {
        r#"console.log(((2n**3n)+1n)*2n);"#,
        ["18n"]
    };

    bigint_from_octal_string => {
        r#"console.log(BigInt("0o10"));"#,
        ["8n"]
    };

    bigint_string_concat_not_add => {
        r#"console.log("n"+String(5n));"#,
        ["n5"]
    };

    bigint_abs_via_condition => {
        r#"const b=-9n; console.log(b<0n?-b:b);"#,
        ["9n"]
    };

    bigint_sort_in_array => {
        r#"console.log([3n,1n,2n].sort((a,b)=>a<b?-1:1).join(","));"#,
        ["1,2,3"]
    };

    bigint_in_map_key => {
        r#"console.log(new Map([[1n,"a"]]).get(1n));"#,
        ["a"]
    };

    bigint_in_set_membership => {
        r#"console.log(new Set([1n,2n]).has(2n));"#,
        ["true"]
    };

    bigint_division_one => {
        r#"console.log((100n/1n).toString());"#,
        ["100"]
    };

    bigint_remainder_equal_dividend => {
        r#"console.log((7n%7n).toString());"#,
        ["0"]
    };

    bigint_power_of_zero => {
        r#"console.log((99n**0n).toString());"#,
        ["1"]
    };

    bigint_zero_power => {
        r#"console.log((0n**5n).toString());"#,
        ["0"]
    };

    bigint_constructor_invalid_string_throws => {
        r#"try{BigInt("abc");}catch(e){console.log(e instanceof SyntaxError);}"#,
        ["true"]
    };
}
