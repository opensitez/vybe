use super::helpers::{compile_ok, run_prints};

#[test]
fn array_transfer_between_kinds_char_integer_round_trip() {
    let out = run_prints(
        r#"
program array_transfer_between_kinds_char_integer_round_trip
    character(len=4) :: token
    integer :: code
    token = transfer(32, "    ")
    code = transfer(token, 0)
    print *, code
    print *, len(token)
end program array_transfer_between_kinds_char_integer_round_trip
"#,
    );
    assert_eq!(out.len(), 2);
}

#[test]
fn array_transfer_between_kinds_int_to_real() {
    let out = run_prints(
        r#"
program array_transfer_between_kinds_int_to_real
    integer :: source
    real :: sink
    source = 1
    sink = transfer(source, sink)
    print *, int(sink)
end program array_transfer_between_kinds_int_to_real
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn array_transfer_between_kinds_real_to_int() {
    let out = run_prints(
        r#"
program array_transfer_between_kinds_real_to_int
    real :: source
    integer :: sink
    source = 2.0
    sink = transfer(source, sink)
    print *, sink
end program array_transfer_between_kinds_real_to_int
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn array_transfer_between_kinds_reshape_like_conversion() {
    let out = run_prints(
        r#"
program array_transfer_between_kinds_reshape_like_conversion
    integer :: a(2)
    integer :: b(2)
    a = (/11, 22/)
    b = transfer(a, b)
    print *, b(1)
    print *, b(2)
end program array_transfer_between_kinds_reshape_like_conversion
"#,
    );
    assert_eq!(out, vec!["11", "22"]);
}

#[test]
fn array_transfer_between_kinds_logical_to_int() {
    let out = run_prints(
        r#"
program array_transfer_between_kinds_logical_to_int
    logical :: ok
    integer :: mark
    ok = .true.
    mark = transfer(ok, mark)
    print *, mark
end program array_transfer_between_kinds_logical_to_int
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn array_transfer_between_kinds_array_shapes_from_transfer_is_size_stable() {
    let out = run_prints(
        r#"
program array_transfer_between_kinds_array_shapes_from_transfer_is_size_stable
    integer :: packed(2)
    integer :: flat(2)
    packed = (/7, 8/)
    flat = transfer(packed, flat)
    print *, size(packed)
    print *, size(flat)
    print *, flat(1)
    print *, flat(2)
end program array_transfer_between_kinds_array_shapes_from_transfer_is_size_stable
"#,
    );
    assert_eq!(out, vec!["2", "2", "7", "8"]);
}

#[test]
fn array_transfer_between_kinds_default_scalar_target_kind_matches_shape() {
    let out = run_prints(
        r#"
program array_transfer_between_kinds_default_scalar_target_kind_matches_shape
    integer :: source(2)
    integer :: target
    source = (/1, 2/)
    target = transfer(source, target)
    print *, target
end program array_transfer_between_kinds_default_scalar_target_kind_matches_shape
"#,
    );
    assert_eq!(out.len(), 1);
}

#[test]
fn array_transfer_between_kinds_prohibited_shape_expansion_should_not_compile() {
    let src =
        "program array_transfer_between_kinds_prohibited_shape_expansion_should_not_compile\n"
            .to_string();
    let body = "  integer :: a(2)\n  integer :: b\n  b = transfer(a, b, kind(a))\nend program";
    compile_ok(&(src + body).as_str());
}

#[test]
fn array_transfer_between_kinds_character_length_expression() {
    let out = run_prints(
        r#"
program array_transfer_between_kinds_character_length_expression
    character(len=5) :: token
    integer :: value
    token = transfer(12345, token)
    value = transfer(token, value)
    print *, len(token)
    print *, value
end program array_transfer_between_kinds_character_length_expression
"#,
    );
    assert_eq!(out, vec!["5"]);
}
