! vybe-test: fortran/random_number_extended/random_seed_get_after_put
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
integer :: s1(2) = [11, 22], s2(2)
call random_seed(put=s1)
call random_seed(get=s2)
if ((merge(1, 0, s2(1) == 11 .and. s2(2) == 22)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, s2(1) == 11 .and. s2(2) == 22), "]"
    stop 1
end if
end program t
