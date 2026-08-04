! vybe-test: fortran/kind_inquiry/precision_array_real
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
real, dimension(3) :: x
x = [1.0, 2.0, 3.0]
if ((precision(x)) /= 53) then
    print *, "FAIL: want [53] got [", precision(x), "]"
    stop 1
end if
end program t
