! vybe-test: fortran/kind_inquiry/range_real_kind_four_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
real(kind=4) :: x = 0.0_4
if ((range(x)) /= 37) then
    print *, "FAIL: want [37] got [", range(x), "]"
    stop 1
end if
end program t
