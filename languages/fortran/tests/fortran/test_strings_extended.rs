use super::helpers::{compile_ok, run_prints};

// ── INDEX — substring position ────────────────────────────────

#[test]
fn index_found() {
    let out = run_prints(
        "program t\ncharacter(len=20) :: s = 'hello world'\nprint *, index(s, 'world')\nend program t\n",
    );
    assert_eq!(out, ["7"]);
}

#[test]
fn index_not_found() {
    let out = run_prints(
        "program t\ncharacter(len=20) :: s = 'hello world'\nprint *, index(s, 'xyz')\nend program t\n",
    );
    assert_eq!(out, ["0"]);
}

#[test]
fn index_from_back() {
    let out = run_prints(
        "program t\ncharacter(len=20) :: s = 'abcabc'\nprint *, index(s, 'bc', .true.)\nend program t\n",
    );
    assert_eq!(out, ["5"]);
}

#[test]
fn count_occurrences_runtime() {
    let out = run_prints(
        "program t\nprint *, count_occurrences('the quick brown fox jumps over the lazy dog', 'the')\ncontains\npure function count_occurrences(s, sub) result(n)\ncharacter(len=*), intent(in) :: s, sub\ninteger :: n, pos, start, lsub\nn = 0\nlsub = len_trim(sub)\nif (lsub == 0) return\nstart = 1\ndo\n    pos = index(s(start:), trim(sub))\n    if (pos == 0) exit\n    n = n + 1\n    start = start + pos + lsub - 1\nend do\nend function count_occurrences\nend program t\n",
    );
    assert_eq!(out, ["2"]);
}

// ── SCAN — find first char in set ────────────────────────────

#[test]
fn scan_found() {
    let out = run_prints(
        "program t\ncharacter(len=10) :: s = 'hello'\nprint *, scan(s, 'aeiou')\nend program t\n",
    );
    assert_eq!(out, ["2"]);
}

#[test]
fn scan_not_found() {
    let out = run_prints(
        "program t\ncharacter(len=10) :: s = 'bcdfg'\nprint *, scan(s, 'aeiou')\nend program t\n",
    );
    assert_eq!(out, ["0"]);
}

#[test]
fn scan_back() {
    let out = run_prints(
        "program t\ncharacter(len=10) :: s = 'hello'\nprint *, scan(s, 'aeiou', .true.)\nend program t\n",
    );
    assert_eq!(out, ["5"]);
}

// ── VERIFY — find first char NOT in set ──────────────────────

#[test]
fn verify_all_in_set() {
    let out = run_prints(
        "program t\ncharacter(len=10) :: s = 'aabbcc'\nprint *, verify(s, 'abc')\nend program t\n",
    );
    assert_eq!(out, ["0"]);
}

#[test]
fn verify_not_in_set() {
    let out = run_prints(
        "program t\ncharacter(len=10) :: s = 'hello'\nprint *, verify(s, 'aeiou')\nend program t\n",
    );
    assert_eq!(out, ["1"]);
}

// ── ADJUSTR — right-adjust ────────────────────────────────────

#[test]
fn adjustr_basic() {
    let out = run_prints(
        "program t\ncharacter(len=10) :: s = 'hello'\ncharacter(len=10) :: r\nr = adjustr(s)\nprint *, len_trim(r)\nend program t\n",
    );
    assert_eq!(out, ["5"]);
}

#[test]
fn adjustl_result() {
    let out = run_prints(
        "program t\ncharacter(len=10) :: s = '  hello'\nprint *, trim(adjustl(s))\nend program t\n",
    );
    assert_eq!(out, ["hello"]);
}

// ── Character slicing ─────────────────────────────────────────

#[test]
fn char_slice_range() {
    let out = run_prints(
        "program t\ncharacter(len=10) :: s = 'abcdefgh'\ncharacter(len=3) :: sub\nsub = s(2:4)\nprint *, sub\nend program t\n",
    );
    assert_eq!(out, ["bcd"]);
}

#[test]
fn char_slice_from_start() {
    let out = run_prints(
        "program t\ncharacter(len=10) :: s = 'hello'\ncharacter(len=3) :: sub\nsub = s(:3)\nprint *, sub\nend program t\n",
    );
    assert_eq!(out, ["hel"]);
}

#[test]
fn char_slice_to_end() {
    let out = run_prints(
        "program t\ncharacter(len=10) :: s = 'hello'\ncharacter(len=5) :: sub\nsub = s(3:)\nprint *, trim(sub)\nend program t\n",
    );
    assert_eq!(out, ["llo"]);
}

#[test]
fn char_slice_assignment_runtime() {
    let out = run_prints(
        "program t\ncharacter(len=5) :: s = 'abcde'\ns(2:4) = 'XYZ'\nprint *, trim(s)\nend program t\n",
    );
    assert_eq!(out, ["aXYZe"]);
}

// ── LEN / LEN_TRIM ───────────────────────────────────────────

#[test]
fn len_padded() {
    let out =
        run_prints("program t\ncharacter(len=10) :: s = 'hi'\nprint *, len(s)\nend program t\n");
    assert_eq!(out, ["10"]);
}

#[test]
fn len_trim_padded() {
    let out = run_prints(
        "program t\ncharacter(len=10) :: s = 'hi'\nprint *, len_trim(s)\nend program t\n",
    );
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

#[test]
fn to_upper_and_to_lower_runtime() {
    let out = run_prints(
        "program t\ncharacter(len=5) :: s = 'AbCdE'\nprint *, to_upper(s)\nprint *, to_lower(s)\nend program t\n",
    );
    assert_eq!(out, ["ABCDE", "abcde"]);
}

#[test]
fn local_function_string_slice_runtime() {
    let out = run_prints(
        "program t\nprint *, trim(str_upper('ab'))\ncontains\npure function str_upper(s) result(u)\ncharacter(len=*), intent(in) :: s\ncharacter(len=len(s)) :: u\ninteger :: i\ndo i = 1, len(s)\n    u(i:i) = s(i:i)\nend do\nend function str_upper\nend program t\n",
    );
    assert_eq!(out, ["ab"]);
}

#[test]
fn shared_str_split_runtime() {
    let out = run_prints(
        "program t\ncharacter(len=256), allocatable :: tokens(:)\ntokens = str_split('alpha:beta:gamma', ':')\nprint *, size(tokens)\nprint *, trim(tokens(1))\nprint *, trim(tokens(3))\nend program t\n",
    );
    assert_eq!(out, ["3", "alpha", "gamma"]);
}

#[test]
fn shared_str_split_direct_index_runtime() {
    let out = run_prints(
        "program t\nprint *, trim(str_split('alpha:beta:gamma', ':')(1))\nprint *, trim(str_split('alpha:beta:gamma', ':')(3))\nend program t\n",
    );
    assert_eq!(out, ["alpha", "gamma"]);
}

#[test]
fn shared_array_join_runtime() {
    let out = run_prints(
        "program t\ncharacter(len=256), allocatable :: tokens(:)\ntokens = str_split('alpha:beta:gamma', ':')\nprint *, trim(array_join(tokens, ' | '))\nend program t\n",
    );
    assert_eq!(out, ["alpha | beta | gamma"]);
}

#[test]
fn shared_array_join_direct_runtime() {
    let out = run_prints(
        "program t\nprint *, trim(array_join(str_split('alpha:beta:gamma', ':'), ' | '))\nend program t\n",
    );
    assert_eq!(out, ["alpha | beta | gamma"]);
}

#[test]
fn shared_str_getcsv_runtime() {
    let out = run_prints(
        "program t\ncharacter(len=256), allocatable :: fields(:)\nfields = str_getcsv('\"Smith, John\",42,\"New York\",\"Engineer, Senior\",95000.50')\nprint *, size(fields)\nprint *, trim(fields(1))\nprint *, trim(fields(4))\nprint *, trim(fields(5))\nend program t\n",
    );
    assert_eq!(out, ["5", "Smith, John", "Engineer, Senior", "95000.50"]);
}

#[test]
fn shared_str_getcsv_direct_runtime() {
    let out = run_prints(
        "program t\nprint *, size(str_getcsv('\"Smith, John\",42,\"New York\",\"Engineer, Senior\",95000.50'))\nend program t\n",
    );
    assert_eq!(out, ["5"]);
}

#[test]
fn str_split_with_empty_fields() {
    let out = run_prints(
        "program t\ncharacter(len=10), allocatable :: tokens(:)\ntokens = str_split('a,,b', ',')\nprint *, size(tokens)\nprint *, trim(tokens(1))\nprint *, len_trim(trim(tokens(2)))\nprint *, trim(tokens(3))\nend program t\n",
    );
    assert_eq!(out, ["3", "a", "0", "b"]);
}

#[test]
fn str_split_multi_char_delimiter_chain() {
    let out = run_prints(
        "program t\ncharacter(len=20), allocatable :: parts(:)\nparts = str_split('x--y--z--', '--')\nprint *, size(parts)\nprint *, trim(parts(1))\nprint *, trim(parts(2))\nprint *, len_trim(trim(parts(4)))\nend program t\n",
    );
    assert_eq!(out, ["4", "x", "y", "0"]);
}

#[test]
fn string_slice_len_trim_chain() {
    let out = run_prints(
        "program t\ncharacter(len=12) :: s\ns = '  trim_me  '\nprint *, len_trim(s(3:))\nprint *, len_trim(s(:4))\nprint *, len_trim(s(3:6))\nend program t\n",
    );
    assert_eq!(out, ["7", "4", "4"]);
}

// ── Lexicographic comparison ──────────────────────────────────

#[test]
fn lge_equal() {
    let out =
        run_prints("program t\nlogical :: b\nb = lge('abc', 'abc')\nprint *, b\nend program t\n");
    assert_eq!(out, ["true"]);
}

#[test]
fn lgt_greater() {
    let out = run_prints("program t\nlogical :: b\nb = lgt('b', 'a')\nprint *, b\nend program t\n");
    assert_eq!(out, ["true"]);
}

#[test]
fn lle_less_equal() {
    let out = run_prints("program t\nlogical :: b\nb = lle('a', 'b')\nprint *, b\nend program t\n");
    assert_eq!(out, ["true"]);
}

#[test]
fn llt_less() {
    let out = run_prints("program t\nlogical :: b\nb = llt('a', 'b')\nprint *, b\nend program t\n");
    assert_eq!(out, ["true"]);
}

// ── String concatenation ─────────────────────────────────────

#[test]
fn concat_result() {
    let out = run_prints(
        "program t\ncharacter(len=10) :: a = 'Hello'\ncharacter(len=10) :: b = ' World'\ncharacter(len=20) :: c\nc = trim(a) // trim(b)\nprint *, trim(c)\nend program t\n",
    );
    assert_eq!(out, ["Hello World"]);
}

#[test]
fn concat_three() {
    compile_ok(
        r#"
program test
    character(len=5) :: a = 'Hello'
    character(len=2) :: b = ', '
    character(len=5) :: c = 'World'
    character(len=15) :: s
    s = trim(a) // trim(b) // trim(c)
    print *, trim(s)
end program test
"#,
    );
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
    let out = run_prints(
        "program t\ncharacter(len=10) :: s = 'hello'\nprint *, trim(s)\nend program t\n",
    );
    assert_eq!(out, ["hello"]);
}

// ── Character comparison ──────────────────────────────────────

#[test]
fn char_compare_eq() {
    let out = run_prints(
        "program t\ncharacter(len=5) :: a = 'hello'\ncharacter(len=5) :: b = 'hello'\nprint *, a == b\nend program t\n",
    );
    assert_eq!(out, ["true"]);
}

#[test]
fn char_compare_ne() {
    let out = run_prints(
        "program t\ncharacter(len=5) :: a = 'hello'\ncharacter(len=5) :: b = 'world'\nprint *, a /= b\nend program t\n",
    );
    assert_eq!(out, ["true"]);
}

#[test]
fn char_compare_lt() {
    let out = run_prints(
        "program t\ncharacter(len=5) :: a = 'apple'\ncharacter(len=6) :: b = 'banana'\nprint *, a < b\nend program t\n",
    );
    assert_eq!(out, ["true"]);
}

// ── Assumed-length character parameters ───────────────────────

#[test]
fn char_assumed_len_arg() {
    let out = run_prints(
        r#"
program test
    call print_it('alice')
    call print_it('bob')
contains
    subroutine print_it(s)
        character(len=*), intent(in) :: s
        print *, trim(s)
    end subroutine
end program test
"#,
    );
    assert_eq!(out, ["alice", "bob"]);
}

#[test]
fn char_len_star_function() {
    let out = run_prints(
        r#"
program test
    call show('hello')
contains
    subroutine show(msg)
        character(len=*), intent(in) :: msg
        print *, trim(msg)
    end subroutine
end program test
"#,
    );
    assert_eq!(out, ["hello"]);
}

// ── String in derived type ────────────────────────────────────

#[test]
fn char_in_derived_type() {
    compile_ok(
        r#"
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
"#,
    );
}
