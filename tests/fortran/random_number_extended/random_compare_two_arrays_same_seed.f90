! vybe-test: fortran/random_number_extended/random_compare_two_arrays_same_seed
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
integer :: seed(1) = [8080]
real :: a(3), b(3)
call random_seed(put=seed)
call random_number(a)
call random_seed(put=seed)
call random_number(b)
if ((merge(1, 0, a(1) == b(1) .and. a(2) == b(2))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, a(1) == b(1) .and. a(2) == b(2)), "]"
    stop 1
end if
end program t
