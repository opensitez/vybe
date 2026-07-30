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

#[test]
fn test_io_nonadvancing_character_modes_chain_and_readback_chars() {
    let out = run_prints(
        r#"
program test_io_nonadvancing_character_modes
    integer :: unit
    character(len=10) :: a, b
    open(newunit=unit, status='scratch', action='readwrite')
    write(unit, '(A)', advance='no') 'x'
    write(unit, '(A)', advance='no') 'y'
    write(unit, '(A)', advance='no') 'z'
    rewind(unit)
    read(unit, '(A1)') a
    read(unit, '(A1)') b
    print *, a
    print *, b
    close(unit)
end program test_io_nonadvancing_character_modes
"#,
    );

    assert_eq!(out, vec!["x", "y"]);
}

#[test]
fn test_io_nonadvancing_character_modes_advance_to_newline_then_more() {
    let out = run_prints(
        r#"
program test_io_nonadvancing_character_modes
    integer :: unit
    character(len=20) :: line
    open(newunit=unit, status='scratch', action='readwrite')
    write(unit, '(A)', advance='no') 'left'
    write(unit, '(A)', advance='yes') ','
    write(unit, '(A)', advance='no') 'right'
    rewind(unit)
    read(unit, '(A)') line
    print *, trim(line)
    close(unit)
end program test_io_nonadvancing_character_modes
"#,
    );

    assert_eq!(out, vec!["left,right"]);
}

#[test]
fn test_io_nonadvancing_character_modes_read_iostat() {
    let out = run_prints(
        r#"
program test_io_nonadvancing_character_modes
    integer :: ios
    integer :: unit
    character(len=2) :: token
    open(newunit=unit, status='scratch', action='readwrite')
    write(unit, '(A)', advance='no') 'q'
    rewind(unit)
    read(unit, '(A2)', iostat=ios) token
    print *, trim(token)
    if (ios /= 0) then
        print *, 1
    else
        print *, 0
    end if
    close(unit)
end program test_io_nonadvancing_character_modes
"#,
    );

    assert_eq!(out, vec!["q", "0"]);
}
