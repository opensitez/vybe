! vybe-test: fortran/recursive_procedures_and_terminators/recursive_procedures_and_terminators_branching_termination
! origin: languages/fortran/tests/fortran/test_recursive_procedures_and_terminators.rs

program recursive_procedures_and_terminators_branching_termination
    if ((sum_until(1, 10)) /= 55) then
    print *, "FAIL: want [55] got [", sum_until(1, 10), "]"
    stop 1
end if
    if ((sum_until(1, 2)) /= 3) then
    print *, "FAIL: want [3] got [", sum_until(1, 2), "]"
    stop 1
end if
contains
    recursive integer function sum_until(start, limit) result(out)
        integer, intent(in) :: start
        integer, intent(in) :: limit
        if (start > limit) then
            out = 0
        else if (start == limit) then
            out = start
        else
            out = start + sum_until(start + 1, limit)
        end if
    end function sum_until
end program recursive_procedures_and_terminators_branching_termination
