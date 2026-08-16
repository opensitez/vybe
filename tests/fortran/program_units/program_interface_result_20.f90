! vybe-test: fortran/program_units/program_interface_result_20
! origin: languages/fortran/tests/fortran/test_program_units.rs
integer function f()
f = 8
end function f
program t
interface
integer function f()
end function f
end interface
if (f() /= 8) then
    print *, "FAIL: want [8] got [", f(), "]"
    stop 1
end if
end program t
