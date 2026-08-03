use super::helpers::run_lua_one;

fn assert_i64_range(
    range: std::ops::RangeInclusive<i32>,
    mut source: impl FnMut(i32) -> String,
    mut expected: impl FnMut(i32) -> i64,
) {
    for n in range {
        assert_eq!(run_lua_one(&source(n)), expected(n).to_string());
    }
}

fn assert_bool_range(
    range: std::ops::RangeInclusive<i32>,
    mut source: impl FnMut(i32) -> String,
    mut expected: impl FnMut(i32) -> bool,
) {
    for n in range {
        assert_eq!(run_lua_one(&source(n)), expected(n).to_string());
    }
}

fn assert_str_range(
    range: std::ops::RangeInclusive<i32>,
    mut source: impl FnMut(i32) -> String,
    mut expected: impl FnMut(i32) -> String,
) {
    for n in range {
        assert_eq!(run_lua_one(&source(n)), expected(n));
    }
}

#[test]
fn lua_matrix_arithmetic_add_ten() {
    assert_i64_range(1..=20, |n| format!("print({n} + 10)"), |n| (n + 10) as i64);
}

#[test]
fn lua_matrix_arithmetic_subtract_ten() {
    assert_i64_range(1..=20, |n| format!("print({n} - 10)"), |n| (n - 10) as i64);
}

#[test]
fn lua_matrix_arithmetic_offset_from_constant() {
    assert_i64_range(1..=20, |n| format!("print(100 - {n})"), |n| (100 - n) as i64);
}

#[test]
fn lua_matrix_arithmetic_multiply_two() {
    assert_i64_range(1..=20, |n| format!("print({n} * 2)"), |n| (n * 2) as i64);
}

#[test]
fn lua_matrix_arithmetic_multiply_three() {
    assert_i64_range(1..=20, |n| format!("print(3 * {n})"), |n| (n * 3) as i64);
}

#[test]
fn lua_matrix_arithmetic_multiply_negative_two() {
    assert_i64_range(1..=20, |n| format!("print(-2 * {n})"), |n| (-2 * n) as i64);
}

#[test]
fn lua_matrix_arithmetic_int_div_two() {
    assert_i64_range(1..=20, |n| format!("print({n} // 2)"), |n| (n / 2) as i64);
}

#[test]
fn lua_matrix_arithmetic_int_div_three() {
    assert_i64_range(1..=20, |n| format!("print({n} // 3)"), |n| (n / 3) as i64);
}

#[test]
fn lua_matrix_arithmetic_neg_div_three() {
    assert_i64_range(1..=20, |n| format!("print(-{n} // 3)"), |n| {
        -((n as f64 / 3.0).ceil() as i64)
    });
}

#[test]
fn lua_matrix_arithmetic_mod_three() {
    assert_i64_range(1..=20, |n| format!("print({n} % 3)"), |n| (n % 3) as i64);
}

#[test]
fn lua_matrix_arithmetic_mod_negative_divisor() {
    assert_i64_range(1..=20, |n| format!("print({n} % -3)"), |n| {
        let rem = n % 3;
        match rem {
            0 => 0,
            1 => -2,
            _ => -1 }
    });
}

#[test]
fn lua_matrix_arithmetic_mod_variable_divisor() {
    assert_i64_range(1..=20, |n| {
        let divisor = n % 5 + 1;
        format!("print({n} % {divisor})")
    }, |n| (n % (n % 5 + 1)) as i64);
}

#[test]
fn lua_matrix_arithmetic_power_two() {
    assert_i64_range(1..=20, |n| {
        let exponent = n % 4;
        format!("print(2 ^ {exponent})")
    }, |n| 1_i64 << (n % 4) as u32);
}

#[test]
fn lua_matrix_arithmetic_power_increment() {
    assert_i64_range(
        1..=20,
        |n| format!("print(2 ^ ({n} + 1))"),
        |n| (1_i64 << (n + 1)) / 2,
    );
}

#[test]
fn lua_matrix_arithmetic_precedence_mul_add() {
    assert_i64_range(1..=20, |n| format!("print({n} * 2 + 3)"), |n| (n * 2 + 3) as i64);
}

#[test]
fn lua_matrix_arithmetic_precedence_add_mul() {
    assert_i64_range(
        1..=20,
        |n| format!("print(({n} + 2) * 3)"),
        |n| ((n + 2) * 3) as i64,
    );
}

#[test]
fn lua_matrix_arithmetic_unary_minus() {
    assert_i64_range(1..=20, |n| format!("print(-{n})"), |n| (-n) as i64);
}

#[test]
fn lua_matrix_arithmetic_double_unary() {
    assert_i64_range(1..=20, |n| format!("print(--{n})"), |n| n as i64);
}

#[test]
fn lua_matrix_arithmetic_square() {
    assert_i64_range(1..=20, |n| format!("print({n} * {n})"), |n| (n as i64) * (n as i64));
}

#[test]
fn lua_matrix_arithmetic_square_minus_input() {
    assert_i64_range(
        1..=20,
        |n| format!("print({n} * {n} - {n})"),
        |n| (n as i64) * (n as i64) - n as i64,
    );
}

#[test]
fn lua_matrix_arithmetic_triangle_sum_formula() {
    assert_i64_range(
        1..=20,
        |n| {
            format!("print(({n} * ({n} + 1)) // 2)")
        },
        |n| ((n as i64) * (n as i64 + 1)) / 2,
    );
}

#[test]
fn lua_matrix_arithmetic_ceil_half_as_floor_formula() {
    assert_i64_range(
        1..=20,
        |n| format!("print(({n} + 1) // 2)"),
        |n| ((n + 1) / 2) as i64,
    );
}

#[test]
fn lua_matrix_arithmetic_mixed_div_expr() {
    assert_i64_range(
        1..=20,
        |n| format!("print(({n} + 2) // 3 * 3)"),
        |n| (((n + 2) / 3) * 3) as i64,
    );
}

#[test]
fn lua_matrix_arithmetic_mod_square_small() {
    assert_i64_range(
        1..=20,
        |n| format!("print(({n} * {n}) % 7)"),
        |n| ((n * n) % 7) as i64,
    );
}

#[test]
fn lua_matrix_arithmetic_scale_by_range() {
    assert_i64_range(
        1..=20,
        |n| format!("print({n} * ({n} + 1) // 2)"),
        |n| ((n * (n + 1) / 2) as i64),
    );
}

#[test]
fn lua_matrix_arithmetic_offset_by_constant_chain() {
    assert_i64_range(1..=20, |n| format!("print(({n} + 7) % 11)"), |n| ((n + 7) % 11) as i64);
}

#[test]
fn lua_matrix_arithmetic_constant_mod_large() {
    assert_i64_range(
        1..=20,
        |n| format!("print(({n} + 13) % 17)"),
        |n| ((n + 13) % 17) as i64,
    );
}

#[test]
fn lua_matrix_bool_is_even() {
    assert_bool_range(1..=20, |n| format!("print(({n} % 2) == 0)"), |n| n % 2 == 0);
}

#[test]
fn lua_matrix_bool_is_odd() {
    assert_bool_range(1..=20, |n| format!("print(({n} % 2) ~= 0)"), |n| n % 2 != 0);
}

#[test]
fn lua_matrix_bool_gt_ten() {
    assert_bool_range(1..=20, |n| format!("print({n} > 10)"), |n| n > 10);
}

#[test]
fn lua_matrix_bool_le_ten() {
    assert_bool_range(1..=20, |n| format!("print({n} <= 10)"), |n| n <= 10);
}

#[test]
fn lua_matrix_bool_between_five_and_fifteen() {
    assert_bool_range(
        1..=20,
        |n| format!("print({n} > 5 and {n} < 15)"),
        |n| n > 5 && n < 15,
    );
}

#[test]
fn lua_matrix_bool_between_or_outside_fifteen() {
    assert_bool_range(
        1..=20,
        |n| format!("print(({n} < 5) or ({n} >= 15))"),
        |n| n < 5 || n >= 15,
    );
}

#[test]
fn lua_matrix_bool_eq_thirteen() {
    assert_bool_range(1..=20, |n| format!("print({n} == 13)"), |n| n == 13);
}

#[test]
fn lua_matrix_bool_eq_one_or_last() {
    assert_bool_range(
        1..=20,
        |n| format!("print(({n} == 1) or ({n} == 20))"),
        |n| n == 1 || n == 20,
    );
}

#[test]
fn lua_matrix_bool_divisible_by_three_or_five() {
    assert_bool_range(
        1..=20,
        |n| format!("print({n} % 3 == 0 or {n} % 5 == 0)"),
        |n| n % 3 == 0 || n % 5 == 0,
    );
}

#[test]
fn lua_matrix_bool_divisible_by_three_and_five() {
    assert_bool_range(
        1..=20,
        |n| format!("print(({n} % 3 == 0) and ({n} % 5 == 0))"),
        |n| n % 3 == 0 && n % 5 == 0,
    );
}

#[test]
fn lua_matrix_bool_not_divisible_by_two_and_five() {
    assert_bool_range(
        1..=20,
        |n| format!("print(not (({n} % 2 == 0) and ({n} % 5 == 0))"),
        |n| !((n % 2 == 0) && (n % 5 == 0)),
    );
}

#[test]
fn lua_matrix_bool_mix_chain() {
    assert_bool_range(
        1..=20,
        |n| {
            format!("print((({n} % 2 == 0) and ({n} % 3 == 0)) or ({n} % 7 == 0)")
        },
        |n| ((n % 2 == 0) && (n % 3 == 0)) || (n % 7 == 0),
    );
}

#[test]
fn lua_matrix_bool_nested_truth_expr() {
    assert_bool_range(
        1..=20,
        |n| {
            format!("print(({n} > 2 and {n} < 12) or ({n} > 18))")
        },
        |n| (n > 2 && n < 12) || (n > 18),
    );
}

#[test]
fn lua_matrix_bool_nested_neg_expr() {
    assert_bool_range(
        1..=20,
        |n| format!("print(not (({n} % 2 == 0) and ({n} % 5 == 0))"),
        |n| !((n % 2 == 0) && (n % 5 == 0)),
    );
}

#[test]
fn lua_matrix_string_concat_with_index() {
    assert_str_range(1..=20, |n| format!("print('v' .. {n})"), |n| format!("v{n}"));
}

#[test]
fn lua_matrix_string_concat_formula() {
    assert_str_range(
        1..=20,
        |n| format!("print({n} .. '-' .. ({n} + 1))"),
        |n| format!("{n}-{}", n + 1),
    );
}

#[test]
fn lua_matrix_string_repeat_length_scaled() {
    assert_i64_range(1..=20, |n| {
        let reps = n % 4 + 1;
        format!("print(#(string.rep('a', {reps} .. '')))")
    }, |n| (n % 4 + 1) as i64);
}

#[test]
fn lua_matrix_string_prefix_len() {
    assert_i64_range(
        1..=20,
        |n| format!("print(#('x' .. string.rep('a', {n} % 4 + 1))"),
        |n| (n % 4 + 2) as i64,
    );
}

#[test]
fn lua_matrix_string_uppercase() {
    assert_str_range(1..=20, |n| format!("print(('x' .. {n}):upper())"), |n| format!("X{n}"));
}

#[test]
fn lua_matrix_string_lowercase() {
    assert_str_range(1..=20, |n| format!("print(('X' .. {n}):lower())"), |n| format!("x{n}"));
}

#[test]
fn lua_matrix_string_reverse() {
    assert_str_range(1..=20, |n| format!("print(({'a', {n}}):concat())"), |n| format!("a{n}"));
}

#[test]
fn lua_matrix_string_byte_first_char() {
    assert_i64_range(
        1..=20,
        |n| format!("print(string.byte('abcdef', ({n} - 1) % 6 + 1) )"),
        |n| ((b'a' + ((n - 1) % 6) as u8) as i64),
    );
}

#[test]
fn lua_matrix_string_byte_from_end() {
    assert_i64_range(
        1..=20,
        |n| format!("print(string.byte('abcdef', -(({n} - 1) % 6 + 1)) )"),
        |n| (b'f' - ((n - 1) % 6) as u8) as i64,
    );
}

#[test]
fn lua_matrix_string_sub_center() {
    assert_str_range(
        1..=20,
        |n| {
            let left = (n % 6) + 1;
            let right = left + 1;
            format!("local n = {left}; print(string.sub('abcdef', n, n + 1))")
        },
        |n| {
            let left = (n % 6) + 1;
            let right = (left + 1) as usize;
            let s = "abcdef";
            s[(left - 1) as usize..right as usize].to_string()
        },
    );
}

#[test]
fn lua_matrix_string_gsub_replace_first() {
    assert_str_range(
        1..=20,
        |n| format!("print(('a' .. {n} .. 'a'):gsub('a', 'b', 1))"),
        |n| format!("b{n}a"),
    );
}

#[test]
fn lua_matrix_string_gsub_replace_all() {
    assert_str_range(
        1..=20,
        |n| format!("print(('a' .. {n} .. 'a'):gsub('a', 'z'))"),
        |n| format!("z{n}z"),
    );
}

#[test]
fn lua_matrix_string_match_digits() {
    assert_str_range(1..=20, |n| format!("print(('v' .. {n}):match('%d+') )"), |n| n.to_string());
}

#[test]
fn lua_matrix_string_match_capture() {
    assert_str_range(1..=20, |n| format!("print(('v' .. {n} .. 'x'):match('v(%d+)x'))"), |n| n.to_string());
}

#[test]
fn lua_matrix_string_find_maybe_pos() {
    assert_i64_range(
        1..=20,
        |n| format!("print(('x' .. {n} .. 'y'):find('y'))"),
        |n| if n < 10 { 3 } else { 4 },
    );
}

#[test]
fn lua_matrix_string_find_with_anchor() {
    assert_i64_range(
        1..=20,
        |n| format!("print(('00' .. {n}):find('^0+'))"),
        |_n| 1,
    );
}

#[test]
fn lua_matrix_string_format_width() {
    assert_str_range(1..=20, |n| format!("print(string.format('%03d', {n}))"), |n| format!("{n:03}"));
}

#[test]
fn lua_matrix_string_format_signed() {
    assert_str_range(
        1..=20,
        |n| format!("print(string.format('%+d', {n}) )"),
        |n| format!("+{n}"),
    );
}

#[test]
fn lua_matrix_string_format_signed_negative() {
    assert_str_range(
        1..=20,
        |n| format!("print(string.format('%+d', -{n}) )"),
        |n| format!("-{n}"),
    );
}

#[test]
fn lua_matrix_table_length_growth() {
    assert_i64_range(
        1..=20,
        |n| {
            format!(
                "local t = {{}}; for i = 1, {n} do t[i] = i end; print(#t)"
            )
        },
        |n| n as i64,
    );
}

#[test]
fn lua_matrix_table_sum_iterative() {
    assert_i64_range(
        1..=20,
        |n| {
            format!("local t = {{}}; for i = 1, {n} do t[i] = i end; local s = 0; for _, v in ipairs(t) do s = s + v end; print(s)")
        },
        |n| ((n as i64) * (n as i64 + 1)) / 2,
    );
}

#[test]
fn lua_matrix_table_even_sum() {
    assert_i64_range(
        1..=20,
        |n| {
            format!(
                "local t = {{}}; for i = 1, {n} do t[i] = i end; local s = 0; for _, v in ipairs(t) do if v % 2 == 0 then s = s + v end end; print(s)"
            )
        },
        |n| {
            let max = if n % 2 == 0 { n } else { n - 1 };
            let pairs = max / 2;
            (pairs * (pairs + 1)) as i64
        },
    );
}

#[test]
fn lua_matrix_table_odd_sum() {
    assert_i64_range(
        1..=20,
        |n| {
            format!(
                "local t = {{}}; for i = 1, {n} do t[i] = i end; local s = 0; for _, v in ipairs(t) do if v % 2 == 1 then s = s + v end end; print(s)"
            )
        },
        |n| {
            let pairs = ((n + 1) / 2) as i64;
            pairs * pairs
        },
    );
}

#[test]
fn lua_matrix_table_insert_tail() {
    assert_i64_range(1..=20, |n| {
        format!("local t = {{1}}; table.insert(t, {n}); print(t[2])")
    }, |n| n as i64);
}

#[test]
fn lua_matrix_table_remove_tail() {
    assert_i64_range(
        1..=20,
        |n| format!("local t = {{1, {n}}}; print(table.remove(t, 2))"),
        |n| n as i64,
    );
}

#[test]
fn lua_matrix_table_concat_default_sep() {
    assert_str_range(1..=20, |n| format!("print(table.concat({{{{1, {n}}}}, ',', 1, 2))"), |n| format!("1,{n}"));
}

#[test]
fn lua_matrix_table_concat_pipe_sep() {
    assert_str_range(
        1..=20,
        |n| format!("print(table.concat({{{{1, 2, {n}}}}, '|'))"),
        |n| format!("1|2|{n}"),
    );
}

#[test]
fn lua_matrix_table_unpack_head_tail() {
    assert_i64_range(
        1..=20,
        |n| {
            format!("local a, b = table.unpack({{{{n}, n + 1, {n} + 2}}}); print(a + b)")
        },
        |n| {
            let a = n;
            let b = n + 1;
            (a + b) as i64
        },
    );
}

#[test]
fn lua_matrix_table_rawset_rawget() {
    assert_i64_range(
        1..=20,
        |n| format!("local t = {{}}; rawset(t, {n}, {n}*3); print(rawget(t, {n}))"),
        |n| (n * 3) as i64,
    );
}

#[test]
fn lua_matrix_table_length_after_insert_remove() {
    assert_i64_range(
        1..=20,
        |n| {
            format!(
                "local t = {{}}; for i = 1, {n} do table.insert(t, i) end; table.remove(t, {n}); table.insert(t, 1, 0); print(#t)"
            )
        },
        |n| n as i64,
    );
}

#[test]
fn lua_matrix_table_nested_field_read() {
    assert_i64_range(
        1..=20,
        |n| format!("local t = {{inner = {{value = {n}}}}; print(t.inner.value)"),
        |n| n as i64,
    );
}

#[test]
fn lua_matrix_table_pairs_key_count() {
    assert_i64_range(
        1..=20,
        |n| format!(
            "local t = {{}}; for i = 1, {n} do t[i] = i end; local c = 0; for _, __ in pairs(t) do c = c + 1 end; print(c)"
        ),
        |n| n as i64,
    );
}

#[test]
fn lua_matrix_table_ipairs_sum() {
    assert_i64_range(
        1..=20,
        |n| {
            format!(
                "local t = {{}}; for i = 1, {n} do t[i] = i * 2 end; local s = 0; for _, v in ipairs(t) do s = s + v end; print(s)"
            )
        },
        |n| {
            let n64 = n as i64;
            n64 * (n64 + 1)
        },
    );
}

#[test]
fn lua_matrix_function_simple_return() {
    assert_i64_range(
        1..=20,
        |n| format!("local f = function(v) return v + 1 end; print(f({n}))"),
        |n| (n + 1) as i64,
    );
}

#[test]
fn lua_matrix_function_multiple_return_sum() {
    assert_i64_range(
        1..=20,
        |n| {
            format!(
                "local f = function(v) return v, v + 1, v + 2 end; local a, b, c = f({n}); print(a + b + c)"
            )
        },
        |n| (3 * n + 3) as i64,
    );
}

#[test]
fn lua_matrix_function_variadic_sum() {
    assert_i64_range(
        1..=20,
        |n| {
            format!(
                "local f = function(...) local total = 0; for _, v in ipairs{{...}} do total = total + v end; return total end; print(f({n}, 1, 2, 3))"
            )
        },
        |n| (n + 6) as i64,
    );
}

#[test]
fn lua_matrix_function_optional_second_arg() {
    assert_i64_range(
        1..=20,
        |n| {
            format!(
                "local f = function(v, fallback) fallback = fallback or 5; return v + fallback end; print(f({n}))"
            )
        },
        |n| (n + 5) as i64,
    );
}

#[test]
fn lua_matrix_function_upvalue_counter() {
    assert_i64_range(
        1..=20,
        |n| {
            format!(
                "local c = 0; local f = function() c = c + 1; return c + {n} end; print(f() + f())"
            )
        },
        |n| (2 * n + 3) as i64,
    );
}

#[test]
fn lua_matrix_function_local_factorial_three() {
    assert_i64_range(
        1..=20,
        |n| format!(
            "local function fact(v) if v <= 1 then return 1 else return v*fact(v-1) end end; print(fact({n} % 3 + 1))"
        ),
        |n| match n % 3 { 0 => 1, 1 => 1, _ => 2 },
    );
}

#[test]
fn lua_matrix_function_tail_like_chain() {
    assert_i64_range(
        1..=20,
        |n| {
            format!(
                "local function double(v) return v * 2 end; local function apply(v) return double(v) + {n} end; print(apply({n}))"
            )
        },
        |n| (3 * n) as i64,
    );
}

#[test]
fn lua_matrix_control_if_nested() {
    assert_i64_range(
        1..=20,
        |n| {
            format!(
                "local v = {n}; if v < 5 then v = 1 elseif v < 10 then v = 2 else v = 3 end; print(v)"
            )
        },
        |n| if n < 5 { 1 } else if n < 10 { 2 } else { 3 },
    );
}

#[test]
fn lua_matrix_control_while_accumulate() {
    assert_i64_range(
        1..=20,
        |n| {
            format!(
                "local total = 0; local i = 1; while i <= {n} do total = total + i; i = i + 1 end; print(total)"
            )
        },
        |n| ((n as i64) * (n as i64 + 1)) / 2,
    );
}

#[test]
fn lua_matrix_control_repeat_until() {
    assert_i64_range(
        1..=20,
        |n| {
            format!(
                "local total = 0; local i = 1; repeat total = total + i; i = i + 1; until i > {n}; print(total)"
            )
        },
        |n| ((n as i64) * (n as i64 + 1)) / 2,
    );
}

#[test]
fn lua_matrix_control_for_numeric_step_two() {
    assert_i64_range(
        1..=20,
        |n| format!(
            "local total = 0; for i = 1, {n}, 2 do total = total + i end; print(total)"
        ),
        |n| {
            if n % 2 == 0 {
                let k = (n / 2) as i64;
                k * k
            } else {
                let k = (n / 2 + 1) as i64;
                k * k
            }
        },
    );
}

#[test]
fn lua_matrix_control_for_numeric_descending() {
    assert_i64_range(
        1..=20,
        |n| format!(
            "local total = 0; for i = {n}, 1, -1 do total = total + 1 end; print(total)"
        ),
        |n| n as i64,
    );
}

#[test]
fn lua_matrix_control_do_block_scope() {
    assert_i64_range(
        1..=20,
        |n| {
            format!(
                "local a = 1; do local a = {n}; print(a) end; print(a)"
            )
        },
        |n| n as i64,
    );
}

#[test]
fn lua_matrix_control_break_after_half() {
    assert_i64_range(
        1..=20,
        |n| {
            format!(
                "local sum = 0; for i=1,{n} do if i > ({n}/2) then break end; sum = sum + i end; print(sum)"
            )
        },
        |n| {
            let half = n / 2;
            (half as i64) * (half as i64 + 1) / 2
        },
    );
}

#[test]
fn lua_matrix_control_generic_for_ipairs() {
    assert_i64_range(
        1..=20,
        |n| {
            format!(
                "local t = {{}}; for i=1,{n} do t[i]=i end; local s=0; for _, v in ipairs(t) do s = s + v end; print(s)"
            )
        },
        |n| ((n as i64) * (n as i64 + 1)) / 2,
    );
}

#[test]
fn lua_matrix_coroutine_resume_value() {
    assert_i64_range(
        1..=20,
        |n| {
            format!(
                "local co = coroutine.create(function(v) coroutine.yield(v + 1) end); local ok, first = coroutine.resume(co, {n}); print(first)"
            )
        },
        |n| (n + 1) as i64,
    );
}

#[test]
fn lua_matrix_coroutine_status_flow() {
    assert_str_range(
        1..=20,
        |_| "local co = coroutine.create(function() end); print(coroutine.status(co))".to_string(),
        |_n| "suspended".to_string(),
    );
}

#[test]
fn lua_matrix_coroutine_wrap_add() {
    assert_i64_range(
        1..=20,
        |n| {
            format!(
                "local co = coroutine.wrap(function(v) return v * 2 end); print(co({n}))"
            )
        },
        |n| (2 * n) as i64,
    );
}

#[test]
fn lua_matrix_metatable_index_fallback() {
    assert_i64_range(
        1..=20,
        |n| {
            format!(
                "local proto = {{value = {n}}}; local obj = setmetatable({{}}, {{ __index = proto }}); print(obj.value)"
            )
        },
        |n| n as i64,
    );
}

#[test]
fn lua_matrix_metatable_newindex_guard() {
    assert_i64_range(
        1..=20,
        |n| {
            format!(
                "local t = {{}}; local mt = {{ __newindex = function(_, k, v) rawset(_, k, v * 2) end }}; setmetatable(t, mt); t.a = {n}; print(t.a)"
            )
        },
        |n| (n * 2) as i64,
    );
}

#[test]
fn lua_matrix_metatable_add_operator() {
    assert_i64_range(
        1..=20,
        |n| {
            format!(
                "local mt = {{ __add = function(l, r) return {n} + l.value + r.value end }}; local a = setmetatable({{ value = {n} }}, mt); local b = setmetatable({{ value = 3 }}, mt); print(a + b)"
            )
        },
        |n| (2 * n + 3) as i64,
    );
}

#[test]
fn lua_matrix_metatable_eq_operator() {
    assert_bool_range(
        1..=20,
        |n| {
            let rhs = n + 1;
            format!(
                "local mt = {{ __eq = function(l, r) return l.value == r.value end }}; local a = setmetatable({{ value = {n} }}, mt); local b = setmetatable({{ value = {rhs} }}, mt); print(a == b)"
            )
        },
        |n| n == n + 1,
    );
}
