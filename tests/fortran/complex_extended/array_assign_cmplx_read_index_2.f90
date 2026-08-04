! vybe-test: fortran/complex_extended/array_assign_cmplx_read_index_2
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: x(3)
x(2) = cmplx(11.0, 12.0)
if ((nint(real(x(2)))) /= 11) then
    print *, "FAIL: want [11] got [", nint(real(x(2))), "]"
    stop 1
end if
if ((nint(aimag(x(2)))) /= 12) then
    print *, "FAIL: want [12] got [", nint(aimag(x(2))), "]"
    stop 1
end if
end program t
