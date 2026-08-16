! vybe-test: fortran/program_units/program_function_05
! origin: languages/fortran/tests/fortran/test_program_units.rs
integer function f()
f = 1
end function f
program t
integer :: f
if (f() /= 1) then
    print *, "FAIL: want [1] got [", f(), "]"
    stop 1
end if
end program t
