use super::helpers::run_prints;

#[test]
fn array_shape_casting_assignments_vector_to_matrix_via_reshape() {
    let out = run_prints(
        r#"
program array_shape_casting_assignments_vector_to_matrix_via_reshape
    integer :: flat(6)
    integer :: matrix(2, 3)
    flat = (/1, 2, 3, 4, 5, 6/)
    matrix = reshape(flat, (/2, 3/))
    print *, matrix(1, 1)
    print *, matrix(2, 3)
    print *, sum(matrix)
end program array_shape_casting_assignments_vector_to_matrix_via_reshape
"#,
    );
    assert_eq!(out, vec!["1", "6", "21"]);
}

#[test]
fn array_shape_casting_assignments_matrix_to_vector_via_reshape() {
    let out = run_prints(
        r#"
program array_shape_casting_assignments_matrix_to_vector_via_reshape
    integer :: matrix(3, 2)
    integer :: flat(6)
    matrix = reshape((/1, 2, 3, 4, 5, 6/), (/3, 2/))
    flat = reshape(matrix, (/6/))
    print *, flat(1)
    print *, flat(6)
    print *, sum(flat)
end program array_shape_casting_assignments_matrix_to_vector_via_reshape
"#,
    );
    assert_eq!(out, vec!["1", "6", "21"]);
}

#[test]
fn array_shape_casting_assignments_scalar_to_ranked_array() {
    let out = run_prints(
        r#"
program array_shape_casting_assignments_scalar_to_ranked_array
    integer :: matrix(2, 2)
    matrix = 4
    print *, matrix(1, 1)
    print *, matrix(2, 2)
    print *, sum(matrix)
end program array_shape_casting_assignments_scalar_to_ranked_array
"#,
    );
    assert_eq!(out, vec!["4", "4", "16"]);
}

#[test]
fn array_shape_casting_assignments_nested_reshape_to_array_of_strings() {
    let out = run_prints(
        r#"
program array_shape_casting_assignments_nested_reshape_to_array_of_strings
    character(len=2) :: packed(2)
    integer :: flat(4)
    flat = (/11, 22, 33, 44/)
    packed = reshape(transfer(flat, (/''/)), (/2/))
    print *, packed(1)
    print *, packed(2)
    print *, len_trim(packed(1))
end program array_shape_casting_assignments_nested_reshape_to_array_of_strings
"#,
    );
    assert_eq!(out, vec!["", "", "0"]);
}

#[test]
fn array_shape_casting_assignments_zero_length_preserved_while_casting() {
    let out = run_prints(
        r#"
program array_shape_casting_assignments_zero_length_preserved_while_casting
    integer :: flat(0)
    integer :: matrix(0, 1)
    matrix = reshape(flat, (/0, 1/))
    print *, size(matrix, 1)
    print *, size(matrix, 2)
end program array_shape_casting_assignments_zero_length_preserved_while_casting
"#,
    );
    assert_eq!(out, vec!["0", "1"]);
}

#[test]
fn array_shape_casting_assignments_shape_function_stability() {
    let out = run_prints(
        r#"
program array_shape_casting_assignments_shape_function_stability
    integer :: source(2, 3)
    integer :: first(6)
    integer :: second(3, 2)
    source = reshape((/1, 2, 3, 4, 5, 6/), (/2, 3/))
    first = reshape(source, (/6/))
    second = reshape(first, (/3, 2/))
    print *, shape(first)(1)
    print *, shape(second)(1)
    print *, shape(second)(2)
    print *, sum(second)
end program array_shape_casting_assignments_shape_function_stability
"#,
    );
    assert_eq!(out, vec!["6", "3", "2", "21"]);
}

#[test]
fn array_shape_casting_assignments_transpose_view_shape() {
    let out = run_prints(
        r#"
program array_shape_casting_assignments_transpose_view_shape
    integer :: a(2, 3)
    integer :: b(3, 2)
    a = reshape((/1, 2, 3, 4, 5, 6/), (/2, 3/))
    b = transpose(a)
    print *, b(1, 1)
    print *, b(3, 2)
    print *, sum(b)
end program array_shape_casting_assignments_transpose_view_shape
"#,
    );
    assert_eq!(out, vec!["1", "6", "21"]);
}

#[test]
fn array_shape_casting_assignments_unpack_with_assumed_shape() {
    let out = run_prints(
        r#"
program array_shape_casting_assignments_unpack_with_assumed_shape
    integer :: a(2, 2)
    integer :: b(4)
    call write_back(a, b)
    print *, b(1)
    print *, b(2)
    print *, b(3)
    print *, b(4)
contains
    subroutine write_back(src, dst)
        integer, intent(in)  :: src(:, :)
        integer, intent(out) :: dst(:)
        dst = reshape(src, (/4/))
    end subroutine write_back
end program array_shape_casting_assignments_unpack_with_assumed_shape
"#,
    );
    assert_eq!(out, vec!["1", "2", "3", "4"]);
}

#[test]
fn array_shape_casting_assignments_pack_unpack_pairings() {
    let out = run_prints(
        r#"
program array_shape_casting_assignments_pack_unpack_pairings
    integer :: src(2, 3)
    integer :: dst(3)
    integer :: packed(2)
    src = reshape((/1, 2, 3, 4, 5, 6/), (/2, 3/))
    packed = reshape(reshape(src, (/6/))(1:2), (/2/))
    dst = (/packed(1), packed(2), 9/)
    print *, dst(1)
    print *, dst(2)
    print *, dst(3)
    print *, sum(dst)
end program array_shape_casting_assignments_pack_unpack_pairings
"#,
    );
    assert_eq!(out, vec!["1", "2", "9", "12"]);
}
