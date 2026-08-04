! vybe-test: fortran/fortran_procedure_quality/procedure_quality_recursive_like_depth
! origin: languages/fortran/tests/fortran/test_fortran_procedure_quality.rs

program procedure_quality_recursive_like_depth
    integer :: answer
    answer = countdown(4)
    if ((answer) /= 10) then
    print *, "FAIL: want [10] got [", answer, "]"
    stop 1
end if

contains
    integer function countdown(n)
        integer, intent(in) :: n
        if (n <= 0) then
            countdown = 0
        else
            countdown = n + countdown(n - 1)
        end if
    end function countdown
end program procedure_quality_recursive_like_depth
