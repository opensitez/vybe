use super::helpers::run_prints;

#[test]
fn test_inquire_file_property_matrix_queries_existence_and_size() {
    let out = run_prints(
        r#"
program test_inquire_file_property_matrix
    logical :: exists
    character(len=256) :: path
    path = 'nonexistent_fortran_probe_file.txt'
    inquire(file=trim(path), exist=exists)
    if (exists) then
        print *, 1
    else
        print *, 0
    end if
end program test_inquire_file_property_matrix
"#,
    );

    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_inquire_file_property_matrix_opened_unit_query() {
    let out = run_prints(
        r#"
program test_inquire_file_property_matrix
    integer :: unit
    logical :: is_open
    open(newunit=unit, file='tmp_inquire_opened_probe.txt', status='replace')
    inquire(unit=unit, opened=is_open)
    print *, is_open
    close(unit)
    print *, 0
end program test_inquire_file_property_matrix
"#,
    );

    assert_eq!(out, vec![".TRUE.", "0"]);
}

#[test]
fn test_inquire_file_property_matrix_file_exists_after_create() {
    let out = run_prints(
        r#"
program test_inquire_file_property_matrix
    logical :: exists
    character(len=256) :: path
    path = 'tmp_inquire_exists_probe.txt'
    open(unit=99, file=trim(path), status='replace')
    close(99)
    inquire(file=trim(path), exist=exists)
    if (exists) then
        print *, 1
    else
        print *, 0
    end if
end program test_inquire_file_property_matrix
"#,
    );

    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_inquire_file_property_matrix_reads_access_mode() {
    let out = run_prints(
        r#"
program test_inquire_file_property_matrix
    logical :: exists
    character(len=16) :: access_mode
    character(len=256) :: path
    path = 'tmp_inquire_access_mode.txt'
    open(unit=99, file=trim(path), status='replace')
    inquire(file=trim(path), exist=exists, access=access_mode)
    print *, exists
    print *, trim(access_mode)
    close(99)
end program test_inquire_file_property_matrix
"#,
    );

    assert_eq!(out, vec![".TRUE.", "SEQUENTIAL"]);
}
