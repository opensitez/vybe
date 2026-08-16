! vybe-test: fortran/program_units/program_pointer_arg_29
! origin: languages/fortran/tests/fortran/test_program_units.rs
program t
integer, pointer :: p
integer, target :: v
v = 0
p => v
call s(p)
if (v /= 6) then
    print *, "FAIL: want [6] got [", v, "]"
    stop 1
end if
contains
subroutine s(a)
integer, pointer :: a
a = 6
end subroutine s
end program t
