use super::helpers::run_prints;

#[test]
fn test_io_nonadvancing_character_modes_write_and_readline_chunks() {
    let out = run_prints(
        r#"
program test_io_nonadvancing_character_modes
    integer :: unit
    character(len=40) :: txt
    integer :: n
    open(newunit=unit, status='scratch', action='readwrite')
    write(unit, '(A)', advance='no') 'A'
    write(unit, '(A)', advance='no') 'B'
    rewind(unit)
    read(unit, '(A)') txt
    n = len_trim(txt)
    print *, n
    close(unit)
end program test_io_nonadvancing_character_modes
"#,
    );

    assert_eq!(out, vec!["2"]);
}
