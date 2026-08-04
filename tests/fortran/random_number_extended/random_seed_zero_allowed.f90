! vybe-test: fortran/random_number_extended/random_seed_zero_allowed
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
integer :: seed(1) = [0]
call random_seed(put=seed)
if ((1) /= 1) then
    print *, "FAIL: want [1] got [", 1, "]"
    stop 1
end if
end program t
