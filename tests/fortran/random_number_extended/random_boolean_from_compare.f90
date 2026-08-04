! vybe-test: fortran/random_number_extended/random_boolean_from_compare
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
real :: r
call random_number(r)
if ((r < 1.0) .neqv. .true.) then
    print *, "FAIL: want [true] got [", r < 1.0, "]"
    stop 1
end if
end program t
