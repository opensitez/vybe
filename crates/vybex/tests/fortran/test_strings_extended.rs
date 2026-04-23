use super::helpers::{compile_ok, run_prints};

// ── INDEX — substring position ────────────────────────────────

#[test]
fn index_found() {
    compile_ok("program t\ncharacter(len=20) :: s = 'hello world'\nprint *, index(s, 'world')\nend program t\n");
}

#[test]
fn index_not_found() {
    compile_ok("program t\ncharacter(len=20) :: s = 'hello world'\nprint *, index(s, 'xyz')\nend program t\n");
}

#[test]
fn index_from_back() {
    compile_ok("program t\ncharacter(len=20) :: s = 'abcabc'\nprint *, index(s, 'bc', .true.)\nend program t\n");
}

// ── SCAN — find first char in set ────────────────────────────

#[test]
fn scan_found() {
    compile_ok("program t\ncharacter(len=10) :: s = 'hello'\nprint *, scan(s, 'aeiou')\nend program t\n");
}

#[test]
fn scan_not_found() {
    compile_ok("program t\ncharacter(len=10) :: s = 'bcdfg'\nprint *, scan(s, 'aeiou')\nend program t\n");
}

#[test]
fn scan_back() {
    compile_ok("program t\ncharacter(len=10) :: s = 'hello'\nprint *, scan(s, 'aeiou', .true.)\nend program t\n");
}

// ── VERIFY — find first char NOT in set ──────────────────────

#[test]
fn verify_all_in_set() {
    compile_ok("program t\ncharacter(len=10) :: s = 'aabbcc'\nprint *, verify(s, 'abc')\nend program t\n");
}

#[test]
fn verify_not_in_set() {
    compile_ok("program t\ncharacter(len=10) :: s = 'hello'\nprint *, verify(s, 'aeiou')\nend program t\n");
}

// ── ADJUSTR — right-adjust ────────────────────────────────────

#[test]
fn adjustr_basic() {
    compile_ok("program t\ncharacter(len=10) :: s = 'hello'\ncharacter(len=10) :: r\nr = adjustr(s)\nprint *, len_trim(r)\nend program t\n");
}

#[test]
fn adjustl_result() {
    let out = run_prints("program t\ncharacter(len=10) :: s = '  hello'\nprint *, trim(adjustl(s))\nend program t\n");
    assert_eq!(out, ["hello"]);
}

// ── Character slicing ─────────────────────────────────────────

#[test]
fn char_slice_range() {
    compile_ok("program t\ncharacter(len=10) :: s = 'abcdefgh'\ncharacter(len=3) :: sub\nsub = s(2:4)\nprint *, sub\nend program t\n");
}

#[test]
fn char_slice_from_start() {
    compile_ok("program t\ncharacter(len=10) :: s = 'hello'\ncharacter(len=3) :: sub\nsub = s(:3)\nprint *, sub\nend program t\n");
}

#[test]
fn char_slice_to_end() {
    compile_ok("program t\ncharacter(len=10) :: s = 'hello'\ncharacter(len=5) :: sub\nsub = s(3:)\nprint *, trim(sub)\nend program t\n");
}

// ── LEN / LEN_TRIM ───────────────────────────────────────────

#[test]
fn len_padded() {
    let out = run_prints("program t\ncharacter(len=10) :: s = 'hi'\nprint *, len(s)\nend program t\n");
    assert_eq!(out, ["10"]);
}

#[test]
fn len_trim_padded() {
    let out = run_prints("program t\ncharacter(len=10) :: s = 'hi'\nprint *, len_trim(s)\nend program t\n");
    assert_eq!(out, ["2"]);
}

// ── IACHAR / ACHAR ───────────────────────────────────────────

#[test]
fn iachar_a() {
    compile_ok("program t\nprint *, iachar('A')\nend program t\n");
}

#[test]
fn achar_65() {
    compile_ok("program t\ncharacter :: c\nc = achar(65)\nprint *, c\nend program t\n");
}

#[test]
fn ichar_a() {
    compile_ok("program t\nprint *, ichar('A')\nend program t\n");
}

#[test]
fn char_from_code() {
    compile_ok("program t\ncharacter :: c\nc = char(72)\nprint *, c\nend program t\n");
}

// ── Lexicographic comparison ──────────────────────────────────

#[test]
fn lge_equal() {
    compile_ok("program t\nlogical :: b\nb = lge('abc', 'abc')\nprint *, b\nend program t\n");
}

#[test]
fn lgt_greater() {
    compile_ok("program t\nlogical :: b\nb = lgt('b', 'a')\nprint *, b\nend program t\n");
}

#[test]
fn lle_less_equal() {
    compile_ok("program t\nlogical :: b\nb = lle('a', 'b')\nprint *, b\nend program t\n");
}

#[test]
fn llt_less() {
    compile_ok("program t\nlogical :: b\nb = llt('a', 'b')\nprint *, b\nend program t\n");
}

// ── String concatenation ─────────────────────────────────────

#[test]
fn concat_result() {
    let out = run_prints("program t\ncharacter(len=10) :: a = 'Hello'\ncharacter(len=10) :: b = ' World'\ncharacter(len=20) :: c\nc = trim(a) // trim(b)\nprint *, trim(c)\nend program t\n");
    assert_eq!(out, ["Hello World"]);
}

#[test]
fn concat_three() {
    compile_ok(r#"
program test
    character(len=5) :: a = 'Hello'
    character(len=2) :: b = ', '
    character(len=5) :: c = 'World'
    character(len=15) :: s
    s = trim(a) // trim(b) // trim(c)
    print *, trim(s)
end program test
"#);
}

// ── REPEAT ───────────────────────────────────────────────────

#[test]
fn repeat_result() {
    let out = run_prints("program t\nprint *, repeat('ab', 3)\nend program t\n");
    assert_eq!(out, ["ababab"]);
}

#[test]
fn repeat_one() {
    let out = run_prints("program t\nprint *, repeat('x', 1)\nend program t\n");
    assert_eq!(out, ["x"]);
}

// ── TRIM ─────────────────────────────────────────────────────

#[test]
fn trim_result() {
    let out = run_prints("program t\ncharacter(len=10) :: s = 'hello'\nprint *, trim(s)\nend program t\n");
    assert_eq!(out, ["hello"]);
}

// ── Character comparison ──────────────────────────────────────

#[test]
fn char_compare_eq() {
    compile_ok("program t\ncharacter(len=5) :: a = 'hello'\ncharacter(len=5) :: b = 'hello'\nprint *, a == b\nend program t\n");
}

#[test]
fn char_compare_ne() {
    compile_ok("program t\ncharacter(len=5) :: a = 'hello'\ncharacter(len=5) :: b = 'world'\nprint *, a /= b\nend program t\n");
}

#[test]
fn char_compare_lt() {
    compile_ok("program t\ncharacter(len=5) :: a = 'apple'\ncharacter(len=5) :: b = 'banana'\nprint *, a < b\nend program t\n");
}

// ── Assumed-length character parameters ───────────────────────

#[test]
fn char_assumed_len_arg() {
    compile_ok(r#"
program test
contains
    subroutine print_it(s)
        character(len=*), intent(in) :: s
        print *, trim(s)
    end subroutine
end program test
"#);
}

#[test]
fn char_len_star_function() {
    compile_ok(r#"
program test
    call show('hello')
contains
    subroutine show(msg)
        character(len=*), intent(in) :: msg
        print *, trim(msg)
    end subroutine
end program test
"#);
}

// ── String in derived type ────────────────────────────────────

#[test]
fn char_in_derived_type() {
    compile_ok(r#"
program test
    type :: Person
        character(len=30) :: name
        integer :: age
    end type Person
    type(Person) :: p
    p%name = 'Alice'
    p%age = 30
    print *, trim(p%name)
end program test
"#);
}
