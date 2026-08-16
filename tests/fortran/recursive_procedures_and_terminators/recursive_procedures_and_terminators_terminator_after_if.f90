! vybe-test: fortran/recursive_procedures_and_terminators/recursive_procedures_and_terminators_terminator_after_if
! origin: languages/fortran/tests/fortran/test_recursive_procedures_and_terminators.rs

program recursive_procedures_and_terminators_terminator_after_if
    integer :: sum
    sum = accumulator(1)
    if ((sum) /= 15) then
    print *, "FAIL: want [15] got [", sum, "]"
    stop 1
end if
contains
    recursive integer function accumulator(n) result(out)
        integer, intent(in) :: n
        if (n >= 6) then
            out = 0
        else
            out = n + accumulator(n + 1)
        end if
    end function accumulator
end program recursive_procedures_and_terminators_terminator_after_if
