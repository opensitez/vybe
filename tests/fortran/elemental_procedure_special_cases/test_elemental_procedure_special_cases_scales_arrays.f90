! vybe-test: fortran/elemental_procedure_special_cases/test_elemental_procedure_special_cases_scales_arrays
! origin: languages/fortran/tests/fortran/test_elemental_procedure_special_cases.rs

program test_elemental_procedure_special_cases
    integer :: values(3)
    integer :: output(3)
    values = (/2, 4, 6/)
    output = double(values)
    if ((output(1)) /= 4) then
    print *, "FAIL: want [4] got [", output(1), "]"
    stop 1
end if
    if ((output(2)) /= 8) then
    print *, "FAIL: want [8] got [", output(2), "]"
    stop 1
end if
    if ((output(3)) /= 12) then
    print *, "FAIL: want [12] got [", output(3), "]"
    stop 1
end if

contains
    elemental function double(x) result(y)
        integer, intent(in) :: x
        integer :: y
        y = x * 2
    end function
end program test_elemental_procedure_special_cases
