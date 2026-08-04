! vybe-test: fortran/random_number_extended/random_array_index_single_element
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
real :: a(1)
call random_number(a(1))
if ((merge(1, 0, a(1) >= 0.0 .and. a(1) < 1.0)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, a(1) >= 0.0 .and. a(1) < 1.0), "]"
    stop 1
end if
end program t
