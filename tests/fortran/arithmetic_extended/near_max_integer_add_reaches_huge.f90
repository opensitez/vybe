! vybe-test: fortran/arithmetic_extended/near_max_integer_add_reaches_huge
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
integer :: x
x = 2147483640 + 7
if ((x) /= 2147483647) then
    print *, "FAIL: want [2147483647] got [", x, "]"
    stop 1
end if
end program t
