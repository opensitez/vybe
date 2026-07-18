use super::helpers::run_prints;

#[test]
fn procedure_result_rename_patterns_integer_scaling_result_name() {
    let out = run_prints(
        r#"
program procedure_result_rename_patterns_integer_scaling_result_name
    integer function scaled(v) result(output)
        integer, intent(in) :: v
        output = v * 3
    end function scaled
    print *, scaled(4)
end program procedure_result_rename_patterns_integer_scaling_result_name
"#,
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn procedure_result_rename_patterns_recursive_result_alias() {
    let out = run_prints(
        r#"
program procedure_result_rename_patterns_recursive_result_alias
    integer function fact(n) result(r)
        integer, intent(in) :: n
        if (n <= 1) then
            r = 1
        else
            r = n * fact(n - 1)
        end if
    end function fact
    print *, fact(4)
end program procedure_result_rename_patterns_recursive_result_alias
"#,
    );
    assert_eq!(out, vec!["24"]);
}

#[test]
fn procedure_result_rename_patterns_character_build_result_name() {
    let out = run_prints(
        r#"
program procedure_result_rename_patterns_character_build_result_name
    character(len=16) function with_suffix(base) result(out)
        character(len=*), intent(in) :: base
        out = trim(base) // '_ok'
    end function with_suffix
    print *, trim(with_suffix('test'))
end program procedure_result_rename_patterns_character_build_result_name
"#,
    );
    assert_eq!(out, vec!["test_ok"]);
}

#[test]
fn procedure_result_rename_patterns_logical_return_name() {
    let out = run_prints(
        r#"
program procedure_result_rename_patterns_logical_return_name
    logical function has_even(v) result(flag)
        integer, intent(in) :: v
        flag = mod(v, 2) == 0
    end function has_even
    print *, has_even(8)
    print *, has_even(9)
end program procedure_result_rename_patterns_logical_return_name
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn procedure_result_rename_patterns_array_shape_from_result() {
    let out = run_prints(
        r#"
program procedure_result_rename_patterns_array_shape_from_result
    integer function pick(i) result(v)
        integer, intent(in) :: i
        if (i < 0) v = 0
        if (i >= 0 .and. i <= 2) v = i * 2
        if (i > 2) v = 99
    end function pick
    print *, pick(-1)
    print *, pick(1)
    print *, pick(4)
end program procedure_result_rename_patterns_array_shape_from_result
"#,
    );
    assert_eq!(out, vec!["0", "2", "99"]);
}

#[test]
fn procedure_result_rename_patterns_result_name_in_intent_block() {
    let out = run_prints(
        r#"
program procedure_result_rename_patterns_result_name_in_intent_block
    integer function accumulate(a, b) result(total)
        integer, intent(in) :: a
        integer, intent(in) :: b
        total = a + b
        if (a > b) total = total + 1
    end function accumulate
    print *, accumulate(2, 5)
    print *, accumulate(9, 1)
end program procedure_result_rename_patterns_result_name_in_intent_block
"#,
    );
    assert_eq!(out, vec!["7", "11"]);
}

#[test]
fn procedure_result_rename_patterns_entry_style_result_label() {
    let out = run_prints(
        r#"
program procedure_result_rename_patterns_entry_style_result_label
    integer function normalize(v) result(value)
        integer, intent(in) :: v
        value = v
        if (v < 0) value = -v
    end function normalize
    print *, normalize(-12)
    print *, normalize(15)
end program procedure_result_rename_patterns_entry_style_result_label
"#,
    );
    assert_eq!(out, vec!["12", "15"]);
}

#[test]
fn procedure_result_rename_patterns_nested_function_result_variable() {
    let out = run_prints(
        r#"
program procedure_result_rename_patterns_nested_function_result_variable
    integer function outer(v) result(total)
        integer, intent(in) :: v
        integer :: helper
        helper = v + 1
        total = helper * 2
    end function outer
    print *, outer(3)
end program procedure_result_rename_patterns_nested_function_result_variable
"#,
    );
    assert_eq!(out, vec!["8"]);
}
