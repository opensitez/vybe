! vybe-test: fortran/recursive_procedures_and_terminators/recursive_procedures_and_terminators_terminator_via_zero_stride_guard
! origin: languages/fortran/tests/fortran/test_recursive_procedures_and_terminators.rs

program t
    if ((countdown(3)) /= 3) then
    print *, "FAIL: want [3] got [", countdown(3), "]"
    stop 1
end if
contains
    recursive integer function countdown(n) result(out)
        integer, intent(in) :: n
        if (n <= 0) then
            out = 0
        else
            out = 1 + countdown(n - 1)
        end if
    end function countdown
end program t
