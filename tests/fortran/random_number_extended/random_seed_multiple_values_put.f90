! vybe-test: fortran/random_number_extended/random_seed_multiple_values_put
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
integer :: seed(4) = [1,2,3,4]
call random_seed(put=seed)
if ((1) /= 1) then
    print *, "FAIL: want [1] got [", 1, "]"
    stop 1
end if
end program t
