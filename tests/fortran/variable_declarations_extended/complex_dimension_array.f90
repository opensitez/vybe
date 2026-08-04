! vybe-test: fortran/variable_declarations_extended/complex_dimension_array
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
complex, dimension(2) :: zs
zs(1) = (1.0, 2.0)
zs(2) = (3.0, 4.0)
if ((nint(real(zs(2)))) /= 3) then
    print *, "FAIL: want [3] got [", nint(real(zs(2))), "]"
    stop 1
end if
if ((nint(aimag(zs(2)))) /= 4) then
    print *, "FAIL: want [4] got [", nint(aimag(zs(2))), "]"
    stop 1
end if
end program t
