! vybe-test: fortran/fortran_procedure_quality/procedure_quality_result_keyword
! origin: languages/fortran/tests/fortran/test_fortran_procedure_quality.rs

program procedure_quality_result_keyword
    real :: ratio
    ratio = safe_ratio(10, 4)
    if (abs((ratio) - 2.5) > 1.0e-6) then
    print *, "FAIL: want [2.5] got [", ratio, "]"
    stop 1
end if

contains
    real function safe_ratio(a, b) result(r)
        integer, intent(in) :: a
        integer, intent(in) :: b
        r = real(a) / real(b)
    end function safe_ratio
end program procedure_quality_result_keyword
