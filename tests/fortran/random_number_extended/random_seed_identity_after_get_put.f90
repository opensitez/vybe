! vybe-test: fortran/random_number_extended/random_seed_identity_after_get_put
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
integer :: a(2), b(2)
call random_seed(get=a)
call random_seed(put=a)
call random_seed(get=b)
if ((merge(1, 0, a(1) == b(1) .and. a(2) == b(2))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, a(1) == b(1) .and. a(2) == b(2)), "]"
    stop 1
end if
end program t
