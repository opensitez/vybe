! vybe-test: fortran/associate_construct_extended/associate_complex_part
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
complex :: z = (3.0, 4.0)
associate (re => real(z), im => aimag(z))
if ((int(re + im)) /= 7) then
    print *, "FAIL: want [7] got [", int(re + im), "]"
    stop 1
end if
end associate
end program t
