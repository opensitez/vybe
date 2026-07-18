use super::helpers::run_prints;

#[test]
fn statement_free_form_legacy_line_continuation_with_ampersand() {
    let out = run_prints(
        r#"
program statement_free_form_legacy_line_continuation_with_ampersand
    integer :: value
    value = 1 + &
            2 + &
            3
    print *, value
end program statement_free_form_legacy_line_continuation_with_ampersand
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn statement_free_form_legacy_fixed_style_spacing() {
    let out = run_prints(
        r#"
program statement_free_form_legacy_fixed_style_spacing
    integer :: i
    integer,parameter :: one = 1
    do i=1,3
        print *, i + one
    end do
end program statement_free_form_legacy_fixed_style_spacing
"#,
    );
    assert_eq!(out, vec!["2", "3", "4"]);
}

#[test]
fn statement_free_form_legacy_old_style_do() {
    let out = run_prints(
        r#"
program statement_free_form_legacy_old_style_do
    integer :: i, s
    s = 0
    do 10 i = 1, 4
        s = s + i
10  continue
    print *, s
end program statement_free_form_legacy_old_style_do
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn statement_free_form_legacy_legacy_comment_in_column() {
    let out = run_prints(
        r#"
program statement_free_form_legacy_legacy_comment_in_column
    integer :: a
    a = 10
    !c   legacy comment form preserved
    print *, a
end program statement_free_form_legacy_legacy_comment_in_column
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn statement_free_form_legacy_logical_if_then_omission() {
    let out = run_prints(
        r#"
program statement_free_form_legacy_logical_if_then_omission
    logical :: done
    done = .false.
    if (done) print *, 'skip'
    if (.not. done) print *, 'run'
end program statement_free_form_legacy_logical_if_then_omission
"#,
    );
    assert_eq!(out, vec!["run"]);
}

#[test]
fn statement_free_form_legacy_legacy_data_block() {
    let out = run_prints(
        r#"
program statement_free_form_legacy_legacy_data_block
    integer :: x, y
    data x /3/ y /4/
    print *, x + y
end program statement_free_form_legacy_legacy_data_block
"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn statement_free_form_legacy_legacy_goto_target_alignment() {
    let out = run_prints(
        r#"
program statement_free_form_legacy_legacy_goto_target_alignment
    integer :: x
    x = 1
    if (x .eq. 1) goto 5
    x = 0
5   print *, x
end program statement_free_form_legacy_legacy_goto_target_alignment
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn statement_free_form_legacy_continue_and_labeled_do() {
    let out = run_prints(
        r#"
program statement_free_form_legacy_continue_and_labeled_do
    integer :: i
    i = 0
    do 100 i = 1, 2
        print *, i
100 continue
end program statement_free_form_legacy_continue_and_labeled_do
"#,
    );
    assert_eq!(out, vec!["1", "2"]);
}
