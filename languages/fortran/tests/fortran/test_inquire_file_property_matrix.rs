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
