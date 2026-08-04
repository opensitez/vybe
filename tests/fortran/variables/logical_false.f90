! vybe-test: fortran/variables/logical_false
! origin: languages/fortran/tests/fortran/test_variables.rs
program t
logical :: flag = .false.
if ((flag) .neqv. .false.) then
    print *, "FAIL: want [false] got [", flag, "]"
    stop 1
end if
end program t
