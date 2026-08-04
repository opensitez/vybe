! vybe-test: fortran/variable_declarations_extended/complex_kind_8_decl
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
complex(kind=8) :: z = (1.0_8, 2.0_8)
if ((nint(real(z))) /= 1) then
    print *, "FAIL: want [1] got [", nint(real(z)), "]"
    stop 1
end if
if ((nint(aimag(z))) /= 2) then
    print *, "FAIL: want [2] got [", nint(aimag(z)), "]"
    stop 1
end if
end program t
