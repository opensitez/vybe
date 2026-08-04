! vybe-test: fortran/recursive_procedures_and_terminators/recursive_procedures_and_terminators_return_accumulator
! origin: languages/fortran/tests/fortran/test_recursive_procedures_and_terminators.rs

program recursive_procedures_and_terminators_return_accumulator
    integer :: total
    total = series(1, 4)
    if ((total) /= 10) then
    print *, "FAIL: want [10] got [", total, "]"
    stop 1
end if
contains
    recursive integer function series(a, b) result(out)
        integer, intent(in) :: a, b
        if (a > b) then
            out = 0
        else
            out = a + series(a + 1, b)
        end if
    end function series
end program recursive_procedures_and_terminators_return_accumulator
