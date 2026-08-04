! vybe-test: fortran/variables/logical_true
! origin: languages/fortran/tests/fortran/test_variables.rs
program t
logical :: flag = .true.
if ((flag) .neqv. .true.) then
    print *, "FAIL: want [true] got [", flag, "]"
    stop 1
end if
end program t
