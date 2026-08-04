! vybe-test: fortran/random_number_extended/random_array_assign_from_previous
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
real :: a(2), b(2)
call random_number(a)
b = a
if ((merge(1, 0, b(1) == a(1) .and. b(2) == a(2))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, b(1) == a(1) .and. b(2) == a(2)), "]"
    stop 1
end if
end program t
