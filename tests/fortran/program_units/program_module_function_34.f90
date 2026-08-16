! vybe-test: fortran/program_units/program_module_function_34
! origin: languages/fortran/tests/fortran/test_program_units.rs
module m
contains
integer function f()
f=1
end function f
end module m
program t
use m
if (f() /= 1) then
    print *, "FAIL: want [1] got [", f(), "]"
    stop 1
end if
end program t
