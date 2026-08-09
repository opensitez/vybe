//! Builtin prototype method coverage — distinct behaviors only.
crate::js_cases! {
    number_tofixed_zero_digits => {
        r#"console.log((1.6).toFixed(0));"#,
        ["2"]
    };

    number_tofixed_two_digits => {
        r#"console.log((1.234).toFixed(2));"#,
        ["1.23"]
    };

    number_tofixed_rounds_up => {
        r#"console.log((1.995).toFixed(2));"#,
        ["2.00"]
    };

    number_tofixed_negative => {
        r#"console.log((-1.234).toFixed(2));"#,
        ["-1.23"]
    };

    number_tofixed_large_integer => {
        r#"console.log((123).toFixed(2));"#,
        ["123.00"]
    };

    number_tofixed_range_error_too_large => {
        r#"try{console.log((1).toFixed(101));}catch(e){console.log(e instanceof RangeError);}"#,
        ["true"]
    };

    number_tofixed_range_error_too_small => {
        r#"try{console.log((1).toFixed(-1));}catch(e){console.log(e instanceof RangeError);}"#,
        ["true"]
    };

    number_tofixed_nan_returns_nan => {
        r#"console.log(Number(NaN).toFixed(2));"#,
        ["NaN"]
    };

    number_tofixed_infinity_returns_infinity => {
        r#"console.log((Infinity).toFixed(2));"#,
        ["Infinity"]
    };

    number_toprecision_significant_digits => {
        r#"console.log((123.456).toPrecision(5));"#,
        ["123.46"]
    };

    number_toprecision_small_number => {
        r#"console.log((0.000123).toPrecision(2));"#,
        ["0.00012"]
    };

    number_toprecision_integer_mode => {
        r#"console.log((42).toPrecision(1));"#,
        ["4e+1"]
    };

    number_toprecision_range_error_low => {
        r#"try{console.log((1).toPrecision(0));}catch(e){console.log(e instanceof RangeError);}"#,
        ["true"]
    };

    number_toprecision_range_error_high => {
        r#"try{console.log((1).toPrecision(101));}catch(e){console.log(e instanceof RangeError);}"#,
        ["true"]
    };

    number_toprecision_nan => {
        r#"console.log(Number(NaN).toPrecision(2));"#,
        ["NaN"]
    };

    number_toexponential_with_fraction => {
        r#"console.log((12345).toExponential(2));"#,
        ["1.23e+4"]
    };

    number_toexponential_small => {
        r#"console.log((0.00123).toExponential(2));"#,
        ["1.23e-3"]
    };

    number_toexponential_zero_fraction => {
        r#"console.log((1).toExponential(0));"#,
        ["1e+0"]
    };

    number_toexponential_range_error => {
        r#"try{console.log((1).toExponential(101));}catch(e){console.log(e instanceof RangeError);}"#,
        ["true"]
    };

    number_valueof_primitive => {
        r#"console.log((7).valueOf());"#,
        ["7"]
    };

    number_valueof_object_wrapper => {
        r#"console.log(new Number(7).valueOf());"#,
        ["7"]
    };

    number_tostring_decimal => {
        r#"console.log((255).toString());"#,
        ["255"]
    };

    number_tostring_hex => {
        r#"console.log((255).toString(16));"#,
        ["ff"]
    };

    number_tostring_binary => {
        r#"console.log((5).toString(2));"#,
        ["101"]
    };

    number_tostring_range_error_base => {
        r#"try{console.log((1).toString(37));}catch(e){console.log(e instanceof RangeError);}"#,
        ["true"]
    };

    boolean_tostring_true => {
        r#"console.log(true.toString());"#,
        ["true"]
    };

    boolean_tostring_false => {
        r#"console.log(false.toString());"#,
        ["false"]
    };

    boolean_valueof_true => {
        r#"console.log(true.valueOf());"#,
        ["true"]
    };

    boolean_valueof_false => {
        r#"console.log(false.valueOf());"#,
        ["false"]
    };

    boolean_object_typeof => {
        r#"console.log(typeof new Boolean(true));"#,
        ["object"]
    };

    number_object_typeof => {
        r#"console.log(typeof new Number(1));"#,
        ["object"]
    };

    boolean_primitive_typeof => {
        r#"console.log(typeof false);"#,
        ["boolean"]
    };

    number_negative_zero_tofixed => {
        r#"console.log((-0).toFixed(0));"#,
        ["0"]
    };

    number_max_value_tostring => {
        r#"console.log(typeof Number.MAX_VALUE.toString());"#,
        ["string"]
    };

    number_min_value_toexponential => {
        r#"console.log(Number.MIN_VALUE.toExponential());"#,
        ["5e-324"]
    };

    number_epsilon_value => {
        r#"console.log(Number.EPSILON > 0);"#,
        ["true"]
    };

    number_parsefloat_whitespace => {
        r#"console.log(Number.parseFloat("  3.14  "));"#,
        ["3.14"]
    };

    number_parseint_hex => {
        r#"console.log(Number.parseInt("0x10"));"#,
        ["16"]
    };

    number_isfinite_on_number => {
        r#"console.log(Number.isFinite(1));"#,
        ["true"]
    };

    number_isinteger_on_float => {
        r#"console.log(Number.isInteger(1.5));"#,
        ["false"]
    };

    number_is_safe_integer_boundary => {
        r#"console.log(Number.isSafeInteger(Number.MAX_SAFE_INTEGER));"#,
        ["true"]
    };

    boolean_in_logical_not => {
        r#"console.log(!true); console.log(!false);"#,
        ["false", "true"]
    };

    boolean_equality_boxed => {
        r#"console.log(new Boolean(true)==true); console.log(new Boolean(true)===true);"#,
        ["true", "false"]
    };

    number_equality_boxed => {
        r#"console.log(new Number(5)==5); console.log(new Number(5)===5);"#,
        ["true", "false"]
    };

    // Node-verified: 4 significant digits of -123.456 → "-123.5".
    number_toprecision_on_negative => {
        r#"console.log((-123.456).toPrecision(4));"#,
        ["-123.5"]
    };

    number_tofixed_on_scientific => {
        r#"console.log((1.23e20).toFixed(0));"#,
        ["123000000000000000000"]
    };

    number_toexponential_negative => {
        r#"console.log((-42).toExponential(1));"#,
        ["-4.2e+1"]
    };

    number_valueof_after_arithmetic => {
        r#"let n=10; n=n+5; console.log(n.valueOf());"#,
        ["15"]
    };

    boolean_valueof_in_condition => {
        r#"const b=false; console.log(b.valueOf()===false);"#,
        ["true"]
    };

    number_tostring_negative_base => {
        r#"console.log((-8).toString(2));"#,
        ["-1000"]
    };

    number_tofixed_trailing_zeros => {
        r#"console.log((2).toFixed(4));"#,
        ["2.0000"]
    };

    number_toprecision_exceeds_digits => {
        r#"console.log((12).toPrecision(10));"#,
        ["12.00000000"]
    };

    // Node-verified: isPrototypeOf on a PRIMITIVE is false (§20.1.3.3
    // — no boxing; primitives are not in any prototype chain).
    boolean_prototype_is_boolean => {
        r#"console.log(Boolean.prototype.isPrototypeOf(true));"#,
        ["false"]
    };

    boolean_prototype_on_object => {
        r#"console.log(Boolean.prototype.isPrototypeOf(new Boolean(false)));"#,
        ["true"]
    };

    // Node-verified: false — primitives are not boxed (§20.1.3.3).
    number_prototype_is_number => {
        r#"console.log(Number.prototype.isPrototypeOf(42));"#,
        ["false"]
    };

    number_prototype_on_object => {
        r#"console.log(Number.prototype.isPrototypeOf(new Number(1)));"#,
        ["true"]
    };

    number_nan_toexponential => {
        r#"console.log(Number(NaN).toExponential(2));"#,
        ["NaN"]
    };

    number_infinity_toprecision => {
        r#"console.log((Infinity).toPrecision(3));"#,
        ["Infinity"]
    };

    boolean_object_coerces_in_addition => {
        r#"console.log(new Boolean(false)+1);"#,
        ["1"]
    };

    number_object_coerces_in_addition => {
        r#"console.log(new Number(2)+3);"#,
        ["5"]
    };

    number_tofixed_called_on_integer_literal => {
        r#"console.log((100).toFixed(1));"#,
        ["100.0"]
    };

    number_toprecision_one_digit => {
        r#"console.log((9.99).toPrecision(1));"#,
        ["1e+1"]
    };

    number_toexponential_default_fraction => {
        r#"console.log((1234).toExponential());"#,
        ["1.234e+3"]
    };

    boolean_not_object_wrapper => {
        r#"console.log(typeof new Boolean(true).valueOf());"#,
        ["boolean"]
    };

    number_not_object_wrapper => {
        r#"console.log(typeof new Number(3).valueOf());"#,
        ["number"]
    };

    number_parseint_radix => {
        r#"console.log(Number.parseInt("10",2));"#,
        ["2"]
    };

    number_parsefloat_invalid_nan => {
        r#"console.log(Number.isNaN(Number.parseFloat("not")));"#,
        ["true"]
    };

    number_isfinite_string_false => {
        r#"console.log(Number.isFinite("1"));"#,
        ["false"]
    };

    number_isnan_string_false => {
        r#"console.log(Number.isNaN("NaN"));"#,
        ["false"]
    };

    boolean_tostring_on_object_wrapper => {
        r#"console.log(Object.prototype.toString.call(new Boolean(true)));"#,
        ["[object Boolean]"]
    };

    number_tostring_on_object_wrapper => {
        r#"console.log(Object.prototype.toString.call(new Number(9)));"#,
        ["[object Number]"]
    };

    number_tofixed_fractional_digits_omitted => {
        r#"console.log((1.005).toFixed());"#,
        ["1"]
    };

    number_toprecision_omitted_uses_full => {
        r#"console.log((42).toPrecision());"#,
        ["42"]
    };

    number_toexponential_fraction_omitted => {
        r#"const s=(1.5).toExponential(); console.log(s.includes("e"));"#,
        ["true"]
    };

    boolean_valueof_returns_primitive_not_object => {
        r#"console.log(new Boolean(true).valueOf()===true);"#,
        ["true"]
    };

    number_valueof_returns_primitive_not_object => {
        r#"console.log(new Number(4).valueOf()===4);"#,
        ["true"]
    };

    number_negative_infinity_tofixed => {
        r#"console.log((-Infinity).toFixed(1));"#,
        ["-Infinity"]
    };

    number_zero_toprecision => {
        r#"console.log((0).toPrecision(1));"#,
        ["0"]
    };

    boolean_and_short_circuit => {
        r#"console.log(false&&true);"#,
        ["false"]
    };

    boolean_or_short_circuit => {
        r#"console.log(true||false);"#,
        ["true"]
    };

    number_tostring_zero_base10 => {
        r#"console.log((0).toString(10));"#,
        ["0"]
    };

    number_tofixed_on_negative_infinity => {
        r#"console.log((-Infinity).toFixed(0));"#,
        ["-Infinity"]
    };

    number_toprecision_on_infinity => {
        r#"console.log((Infinity).toPrecision(5));"#,
        ["Infinity"]
    };

    boolean_double_negation => {
        r#"console.log(!!false);"#,
        ["false"]
    };

    number_unary_plus_on_string => {
        r#"console.log(+ "7");"#,
        ["7"]
    };

    number_isinteger_on_bigint_false => {
        r#"console.log(Number.isInteger(1n));"#,
        ["false"]
    };

    number_max_safe_plus_one_not_safe => {
        r#"console.log(Number.isSafeInteger(Number.MAX_SAFE_INTEGER+1));"#,
        ["false"]
    };

    boolean_tostring_on_non_boolean_throws_typeerror => {
        r#"try{Boolean.prototype.toString.call(123);}catch(e){console.log(e instanceof TypeError);}"#,
        ["true"]
    };

}
