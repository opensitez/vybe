use super::helpers::compile_ok;

// ── Declaration ───────────────────────────────────────────────

#[test]
fn char_array_decl() {
    compile_ok(
        r#"
program test
    character(len=10) :: names(5)
    names(1) = 'Alice'
    print *, trim(names(1))
end program test
"#,
    );
}

#[test]
fn char_array_fixed_len() {
    compile_ok(
        r#"
program test
    character(len=1) :: letters(5) = ['a', 'b', 'c', 'd', 'e']
    print *, letters(3)
end program test
"#,
    );
}

#[test]
fn char_array_len20() {
    compile_ok(
        r#"
program test
    character(len=20) :: words(4)
    words(1) = 'Fortran'
    words(2) = 'is'
    words(3) = 'still'
    words(4) = 'relevant'
    print *, trim(words(1)), ' ', trim(words(2))
end program test
"#,
    );
}

// ── Initialization ────────────────────────────────────────────

#[test]
fn char_array_data_init() {
    compile_ok(
        r#"
program test
    character(len=5) :: fruits(3)
    data fruits /'apple', 'mango', 'grape'/
    print *, trim(fruits(2))
end program test
"#,
    );
}

#[test]
fn char_array_loop_init() {
    compile_ok(
        r#"
program test
    character(len=5) :: a(3)
    integer :: i
    a(1) = 'one  '
    a(2) = 'two  '
    a(3) = 'three'
    do i = 1, 3
        print *, trim(a(i))
    end do
end program test
"#,
    );
}

#[test]
fn char_array_param_init() {
    compile_ok(
        r#"
program test
    character(len=3), parameter :: days(7) = &
        ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
    print *, days(1)
    print *, days(5)
end program test
"#,
    );
}

// ── Element access and modification ──────────────────────────

#[test]
fn char_array_element_assign() {
    compile_ok(
        r#"
program test
    character(len=10) :: arr(3)
    arr(1) = 'hello'
    arr(2) = 'world'
    arr(3) = 'foo'
    arr(2) = 'fortran'
    print *, trim(arr(2))
end program test
"#,
    );
}

#[test]
fn char_array_substring_element() {
    compile_ok(
        r#"
program test
    character(len=10) :: arr(3)
    arr(1) = 'Hello World'
    print *, arr(1)(1:5)
end program test
"#,
    );
}

#[test]
fn char_array_substring_assign() {
    compile_ok(
        r#"
program test
    character(len=10) :: arr(2)
    arr(1) = 'XXXXXXXXXX'
    arr(1)(1:5) = 'Hello'
    print *, trim(arr(1))
end program test
"#,
    );
}

// ── Intrinsics on array elements ──────────────────────────────

#[test]
fn char_array_trim() {
    compile_ok(
        r#"
program test
    character(len=20) :: arr(3)
    arr(1) = 'hello     '
    arr(2) = 'world  '
    arr(3) = 'foo'
    print *, len_trim(arr(1))
    print *, len_trim(arr(2))
end program test
"#,
    );
}

#[test]
fn char_array_adjustl() {
    compile_ok(
        r#"
program test
    character(len=10) :: arr(2)
    arr(1) = '   hi'
    arr(2) = ' world'
    print *, trim(adjustl(arr(1)))
    print *, trim(adjustl(arr(2)))
end program test
"#,
    );
}

#[test]
fn char_array_index() {
    compile_ok(
        r#"
program test
    character(len=20) :: arr(3)
    arr(1) = 'hello world'
    arr(2) = 'fortran 90'
    arr(3) = 'no match here'
    print *, index(arr(1), 'world')
    print *, index(arr(3), 'xyz')
end program test
"#,
    );
}

#[test]
fn char_array_len() {
    compile_ok(
        r#"
program test
    character(len=15) :: arr(3)
    arr(1) = 'short'
    print *, len(arr(1))
    print *, len_trim(arr(1))
end program test
"#,
    );
}

// ── Comparison of array elements ─────────────────────────────

#[test]
fn char_array_compare_elements() {
    compile_ok(
        r#"
program test
    character(len=5) :: arr(3) = ['apple', 'mango', 'grape']
    print *, arr(1) < arr(2)
    print *, arr(2) > arr(3)
    print *, arr(1) == arr(1)
end program test
"#,
    );
}

#[test]
fn char_array_lge_llt() {
    compile_ok(
        r#"
program test
    character(len=5) :: arr(2) = ['abc  ', 'xyz  ']
    print *, llt(arr(1), arr(2))
    print *, lge(arr(2), arr(1))
end program test
"#,
    );
}

// ── Char array in derived type ────────────────────────────────

#[test]
fn char_array_in_type() {
    compile_ok(
        r#"
program test
    type :: Person
        character(len=20) :: name
        integer :: tags(3)
        character(len=10) :: labels(3)
    end type Person
    type(Person) :: p
    p%name = 'Alice'
    p%labels(1) = 'engineer'
    p%labels(2) = 'pilot'
    p%labels(3) = 'runner'
    print *, trim(p%name)
    print *, trim(p%labels(2))
end program test
"#,
    );
}

#[test]
fn array_of_types_with_char() {
    compile_ok(
        r#"
program test
    type :: Tag
        character(len=20) :: name
        integer :: value
    end type Tag
    type(Tag) :: tags(3)
    tags(1)%name = 'alpha'
    tags(1)%value = 1
    tags(2)%name = 'beta'
    tags(2)%value = 2
    tags(3)%name = 'gamma'
    tags(3)%value = 3
    print *, trim(tags(2)%name)
end program test
"#,
    );
}

// ── Char array subroutine/function arguments ──────────────────

#[test]
fn char_array_subroutine_arg() {
    compile_ok(
        r#"
program test
    character(len=10) :: names(4)
    names = ['Alice     ', 'Bob       ', 'Charlie   ', 'Diana     ']
    call print_names(names)
contains
    subroutine print_names(arr)
        character(len=*), intent(in) :: arr(:)
        integer :: i
        do i = 1, size(arr)
            print *, trim(arr(i))
        end do
    end subroutine print_names
end program test
"#,
    );
}

#[test]
fn char_array_function_result() {
    compile_ok(
        r#"
program test
    character(len=5) :: words(3)
    integer :: longest
    words = ['hi   ', 'hello', 'hey  ']
    longest = find_longest(words)
    print *, longest
contains
    function find_longest(arr) result(maxlen)
        character(len=*), intent(in) :: arr(:)
        integer :: maxlen, i
        maxlen = 0
        do i = 1, size(arr)
            maxlen = max(maxlen, len_trim(arr(i)))
        end do
    end function find_longest
end program test
"#,
    );
}

// ── Allocatable char arrays ───────────────────────────────────

#[test]
fn char_array_allocatable() {
    compile_ok(
        r#"
program test
    character(len=20), allocatable :: arr(:)
    allocate(arr(5))
    arr(1) = 'first'
    arr(5) = 'last'
    print *, trim(arr(1))
    print *, trim(arr(5))
    deallocate(arr)
end program test
"#,
    );
}

#[test]
fn char_array_allocatable_2d() {
    compile_ok(
        r#"
program test
    character(len=10), allocatable :: grid(:,:)
    allocate(grid(3,3))
    grid(1,1) = 'top-left'
    grid(3,3) = 'bot-right'
    print *, trim(grid(1,1))
    deallocate(grid)
end program test
"#,
    );
}

// ── WHERE on char arrays ──────────────────────────────────────

#[test]
fn where_on_char_array() {
    compile_ok(
        r#"
program test
    character(len=5) :: src(4) = ['apple', 'mango', 'grape', 'pear ']
    character(len=5) :: dst(4)
    dst = '     '
    where (src /= 'mango')
        dst = src
    end where
    print *, trim(dst(1))
    print *, trim(dst(2))
end program test
"#,
    );
}

// ── Sorting a char array (simple bubble sort) ─────────────────

#[test]
fn char_array_sort_bubble() {
    compile_ok(
        r#"
program test
    character(len=5) :: arr(4) = ['delta', 'alpha', 'gamma', 'beta ']
    character(len=5) :: tmp
    integer :: i, j
    do i = 1, 3
        do j = 1, 4 - i
            if (arr(j) > arr(j+1)) then
                tmp = arr(j)
                arr(j) = arr(j+1)
                arr(j+1) = tmp
            end if
        end do
    end do
    print *, trim(arr(1))
    print *, trim(arr(4))
end program test
"#,
    );
}
