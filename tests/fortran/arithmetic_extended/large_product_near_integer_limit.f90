! vybe-test: fortran/arithmetic_extended/large_product_near_integer_limit
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
integer :: x
x = 46340 * 46340
if ((x) /= 2147395600) then
    print *, "FAIL: want [2147395600] got [", x, "]"
    stop 1
end if
end program t
