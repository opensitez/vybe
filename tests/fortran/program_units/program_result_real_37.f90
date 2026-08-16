! vybe-test: fortran/program_units/program_result_real_37
! origin: languages/fortran/tests/fortran/test_program_units.rs
program t
if (abs(f() - 1.0) > 1.0e-6) then
    print *, "FAIL: want [1.0] got [", f(), "]"
    stop 1
end if
contains
real function f() result(r)
r = 1.0
end function f
end program t
