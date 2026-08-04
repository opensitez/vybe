! vybe-test: fortran/basics/logical_variable
! origin: languages/fortran/tests/fortran/test_basics.rs

program test
    logical :: flag
    flag = .true.
    if ((flag) .neqv. .true.) then
    print *, "FAIL: want [true] got [", flag, "]"
    stop 1
end if
end program test
