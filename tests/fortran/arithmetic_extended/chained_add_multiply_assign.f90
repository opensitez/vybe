! vybe-test: fortran/arithmetic_extended/chained_add_multiply_assign
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
integer :: x
x = 1 + 2 ** 3 * 4
if ((x) /= 33) then
    print *, "FAIL: want [33] got [", x, "]"
    stop 1
end if
end program t
