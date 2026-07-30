use super::helpers::run_prints;

#[test]
fn substring_operations_01_slice_start_end() {
    let out = run_prints(
        "program p\ncharacter(len=5) :: s='hello'\nprint *, trim(s(1:2))\nend program p\n",
    );
    assert_eq!(out, vec!["he"]);
}

#[test]
fn substring_operations_02_slice_middle() {
    let out = run_prints(
        "program p\ncharacter(len=5) :: s='hello'\nprint *, trim(s(2:4))\nend program p\n",
    );
    assert_eq!(out, vec!["ell"]);
}

#[test]
fn substring_operations_03_slice_to_upper_default() {
    let out = run_prints(
        "program p\ncharacter(len=5) :: s='hello'\nprint *, trim(s(:3))\nend program p\n",
    );
    assert_eq!(out, vec!["hel"]);
}

#[test]
fn substring_operations_04_slice_from_lower_default() {
    let out = run_prints(
        "program p\ncharacter(len=5) :: s='hello'\nprint *, trim(s(3:))\nend program p\n",
    );
    assert_eq!(out, vec!["llo"]);
}

#[test]
fn substring_operations_05_assignment_mid_section() {
    let out = run_prints(
        "program p\ncharacter(len=6) :: s='abcdef'\ns(2:3)='ZZ'\nprint *, trim(s)\nend program p\n",
    );
    assert_eq!(out, vec!["aZZdef"]);
}

#[test]
fn substring_operations_06_assignment_head_character() {
    let out = run_prints(
        "program p\ncharacter(len=6) :: s='abcdef'\ns(1:1)='X'\nprint *, trim(s)\nend program p\n",
    );
    assert_eq!(out, vec!["Xbcdef"]);
}

#[test]
fn substring_operations_07_single_char_index() {
    let out = run_prints(
        "program p\ncharacter(len=6) :: s='abcdef'\nprint *, trim(s(6:6))\nend program p\n",
    );
    assert_eq!(out, vec!["f"]);
}

#[test]
fn substring_operations_08_len_of_slice() {
    let out = run_prints(
        "program p\ncharacter(len=6) :: s='abcdef'\nprint *, len(s(2:5))\nend program p\n",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn substring_operations_09_index_into_slice() {
    let out = run_prints(
        "program p\ncharacter(len=6) :: s='abcdef'\nprint *, index(s(2:), 'de')\nend program p\n",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn substring_operations_10_scan_into_slice() {
    let out = run_prints(
        "program p\ncharacter(len=6) :: s='abcdef'\nprint *, scan(s(2:5), 'd')\nend program p\n",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn substring_operations_runtime_sections_basic() {
    let out = run_prints(
        r#"
program p
character(len=6) :: s='abcdef'
print *, trim(s(1:2))
print *, trim(s(2:))
print *, trim(s(:4))
print *, trim(s(3:3))
end program p
"#,
    );
    assert_eq!(out, vec!["ab", "bcdef", "abcd", "c"]);
}

#[test]
fn substring_operations_runtime_assignments_from_substrings() {
    let out = run_prints(
        r#"
program p
character(len=7) :: s='abcdef'
s(3:5) = s(1:3)
print *, trim(s)
end program p
"#,
    );
    assert_eq!(out, vec!["ababcef"]);
}

#[test]
fn substring_operations_runtime_variable_bounds() {
    let out = run_prints(
        r#"
program p
character(len=12) :: text
integer :: i
integer :: j
text = 'fortran-lang'
i = 2
j = 5
print *, trim(text(i:j))
print *, len(text(i:j))
end program p
"#,
    );
    assert_eq!(out, vec!["orta", "4"]);
}

#[test]
fn substring_operations_runtime_chain_in_expression() {
    let out = run_prints(
        r#"
program p
character(len=20) :: msg
msg = trim('pre:' // 'fortran' // '-' // 'lang')
print *, trim(msg(1:4))
print *, trim(msg(6:10))
print *, trim(msg(12:15))
end program p
"#,
    );
    assert_eq!(out, vec!["pre:", "ortra", "lang"]);
}

#[test]
fn substring_operations_runtime_bounds_with_open_end_variables() {
    let out = run_prints(
        r#"
program p
character(len=6) :: s
integer :: start_idx
s = 'fortran'
start_idx = 3
print *, trim(s(start_idx:))
print *, trim(s(1:start_idx))
end program p
"#,
    );
    assert_eq!(out, vec!["tran", "for"]);
}

#[test]
fn substring_operations_runtime_inplace_overlap_copy() {
    let out = run_prints(
        r#"
program p
character(len=8) :: s
s = 'abcdefgh'
s(3:6) = s(4:7)
print *, trim(s)
end program p
"#,
    );
    assert_eq!(out, vec!["abdefghh"]);
}
