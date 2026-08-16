! vybe-test: fortran/program_units/program_result_complex_38
! origin: languages/fortran/tests/fortran/test_program_units.rs
program t
complex :: z
z = f()
if (abs(real(z) - 1.0) > 1.0e-6) then
    print *, "FAIL: want [1.0] got [", real(z), "]"
    stop 1
end if
if (abs(aimag(z) - 2.0) > 1.0e-6) then
    print *, "FAIL: want [2.0] got [", aimag(z), "]"
    stop 1
end if
contains
complex function f() result(r)
r = (1.0,2.0)
end function f
end program t
