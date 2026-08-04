! vybe-test: fortran/associate_construct_extended/associate_mixed_int_real_expr
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: i = 5
real :: r = 2.5
associate (mix => real(i) + r)
if ((int(mix)) /= 7) then
    print *, "FAIL: want [7] got [", int(mix), "]"
    stop 1
end if
end associate
end program t
