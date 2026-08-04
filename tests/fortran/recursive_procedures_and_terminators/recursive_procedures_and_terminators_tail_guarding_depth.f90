! vybe-test: fortran/recursive_procedures_and_terminators/recursive_procedures_and_terminators_tail_guarding_depth
! origin: languages/fortran/tests/fortran/test_recursive_procedures_and_terminators.rs

program recursive_procedures_and_terminators_tail_guarding_depth
    if ((depth_walk(4)) /= 2) then
    print *, "FAIL: want [2] got [", depth_walk(4), "]"
    stop 1
end if
contains
    recursive integer function depth_walk(n) result(out)
        integer, intent(in) :: n
        if (n <= 0) then
            out = 0
        else
            out = depth_walk(n - 2) + 1
        end if
    end function depth_walk
end program recursive_procedures_and_terminators_tail_guarding_depth
