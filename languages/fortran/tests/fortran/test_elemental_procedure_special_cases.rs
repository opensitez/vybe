use super::helpers::run_prints;

#[test]
fn test_elemental_procedure_special_cases_scales_arrays() {
    let out = run_prints(
        r#"
program test_elemental_procedure_special_cases
    integer :: values(3)
    integer :: output(3)
    values = (/2, 4, 6/)
    output = double(values)
    print *, output(1)
    print *, output(2)
    print *, output(3)

contains
    elemental function double(x) result(y)
        integer, intent(in) :: x
        integer :: y
        y = x * 2
    end function
end program test_elemental_procedure_special_cases
"#,
    );

    assert_eq!(out, vec!["4", "8", "12"]);
}
