! vybe-test: fortran/program_units/program_allocatable_arg_30
! origin: languages/fortran/tests/fortran/test_program_units.rs
program t
integer, allocatable :: v(:)
allocate(v(2))
v = 1
call s(v)
if (size(v) /= 4) then
    print *, "FAIL: want [4] got [", size(v), "]"
    stop 1
end if
if (sum(v) /= 12) then
    print *, "FAIL: want [12] got [", sum(v), "]"
    stop 1
end if
contains
subroutine s(a)
integer, allocatable :: a(:)
deallocate(a)
allocate(a(4))
a = 3
end subroutine s
end program t
